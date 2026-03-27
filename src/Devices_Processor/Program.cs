using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;

namespace Devices_Processor
{
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length < 1 || string.IsNullOrWhiteSpace(args[0]))
            {
                PrintUsage();
                return 1;
            }

            var inputFolder = args[0];

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"[Status] Input folder does not exist: {inputFolder}");
                return 1;
            }

            Console.WriteLine($"[Status] Input folder: {inputFolder}");
            Console.WriteLine("[Status] Scanning group folders for exported JSON files...");

            var groups = FindGroups(inputFolder);
            var totalJsonCount = groups.Sum(group => group.JsonCount);
            var groupsWithJson = groups.Count(group => group.JsonCount > 0);

            Console.WriteLine($"[Status] Groups found: {groups.Count}");
            Console.WriteLine($"[Status] Groups with JSON files: {groupsWithJson}");

            if (totalJsonCount == 0)
            {
                Console.WriteLine("[Status] Nothing to process.");
                Console.WriteLine("[Summary] Total groups: 0; Total JSON files: 0; Excel files created: 0; Excel files failed: 0");
                return 0;
            }

            var outputRoot = GetOutputRoot(inputFolder);
            var outputBaseFolder = Path.Combine(outputRoot, "collected_excels");
            Directory.CreateDirectory(outputBaseFolder);
            Console.WriteLine($"[Status] Excel output root: {outputBaseFolder}");
            var accessDbPath = Path.Combine(outputRoot, "Devices.accdb");
            EnsureAccessDatabase(accessDbPath);
            Console.WriteLine($"[Status] Access DB: {accessDbPath}");

            var imagesFolder = Path.Combine(outputRoot, "downloaded_images");

            var excelCreated = 0;
            var excelFailed = 0;
            var processedJsonCount = 0;

            using (var metadataService = new SiemensMetadataService(imagesFolder))
            {
                for (var groupIndex = 0; groupIndex < groups.Count; groupIndex++)
                {
                    var group = groups[groupIndex];
                    Console.WriteLine();
                    Console.WriteLine(new string('=', 88));
                    Console.WriteLine($"[Group] {groupIndex + 1}/{groups.Count}: {group.Name}");
                    Console.WriteLine($"[Status] JSON files found: {group.JsonCount}");

                    if (group.JsonCount == 0)
                    {
                        Console.WriteLine("[Status] Group has no JSON files. Skipping.");
                        Console.WriteLine(new string('=', 88));
                        continue;
                    }

                    var groupOutputFolder = Path.Combine(outputBaseFolder, group.Name);
                    Directory.CreateDirectory(groupOutputFolder);
                    Console.WriteLine($"[Status] Group output folder: {groupOutputFolder}");

                    for (var projectIndex = 0; projectIndex < group.JsonPaths.Count; projectIndex++)
                    {
                        var jsonPath = group.JsonPaths[projectIndex];
                        var projectFileName = Path.GetFileName(jsonPath);
                        var projectName = Path.GetFileNameWithoutExtension(projectFileName);
                        var outputPath = Path.Combine(groupOutputFolder, projectName + ".xlsx");
                        Console.WriteLine();
                        Console.WriteLine($"[Status] JSON progress: {processedJsonCount + 1}/{totalJsonCount}");
                        Console.WriteLine($"[Status] Project {projectIndex + 1}/{group.JsonPaths.Count}: {projectFileName}");

                        try
                        {
                            var rows = LoadDeviceRows(jsonPath);
                            var recordTimestamp = DateTime.Now;

                            var dbRows = SaveProjectExcel(outputPath, group.Name, projectName, rows, recordTimestamp, metadataService);
                            var accessTableName = SanitizeAccessTableName(projectName);
                            UpsertAccessTable(accessDbPath, accessTableName, dbRows);

                            Console.WriteLine($"[Result] JSON processed: {projectFileName}");
                            Console.WriteLine($"[Result] Rows loaded: {rows.Count}");
                            Console.WriteLine($"[Result] Excel saved: {outputPath}");
                            Console.WriteLine($"[Result] Access table replaced: {accessTableName}");
                            MoveJsonToProcessed(jsonPath);
                            excelCreated++;
                            processedJsonCount++;
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"[Status] Failed to process JSON: {projectFileName}");
                            PrintException(exception);
                            excelFailed++;
                            processedJsonCount++;
                        }
                    }

                    Console.WriteLine(new string('=', 88));
                }
            }

            Console.WriteLine();
            Console.WriteLine("[Status] Processing finished.");
            Console.WriteLine($"[Summary] Total groups: {groups.Count}; Total JSON files: {totalJsonCount}; Excel files created: {excelCreated}; Excel files failed: {excelFailed}");
            return excelFailed == 0 ? 0 : 1;
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  Devices_Processor.exe \"C:\\path\\to\\collected_components\"");
        }

        private static string GetOutputRoot(string inputFolder)
        {
            var inputFullPath = Path.GetFullPath(inputFolder);
            var inputName = new DirectoryInfo(inputFullPath).Name;
            var parent = Directory.GetParent(inputFullPath);

            if (string.Equals(inputName, "collected_components", StringComparison.OrdinalIgnoreCase) && parent != null)
            {
                return parent.FullName;
            }

            return inputFullPath;
        }

        private static List<GroupInfo> FindGroups(string inputFolder)
        {
            var groups = new List<GroupInfo>();

            var directories = Directory.EnumerateDirectories(inputFolder)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var directory in directories)
            {
                var jsonPaths = Directory.EnumerateFiles(directory, "*.json", SearchOption.TopDirectoryOnly)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                groups.Add(new GroupInfo
                {
                    Name = Path.GetFileName(directory),
                    Path = directory,
                    JsonPaths = jsonPaths,
                    JsonCount = jsonPaths.Count
                });
            }

            return groups;
        }

        private static List<DeviceExportRow> LoadDeviceRows(string jsonPath)
        {
            var serializer = new DataContractJsonSerializer(typeof(List<DeviceExportRow>));

            using (var stream = File.OpenRead(jsonPath))
            {
                var obj = serializer.ReadObject(stream);
                var rows = obj as List<DeviceExportRow>;
                return rows ?? new List<DeviceExportRow>();
            }
        }

        private static void MoveJsonToProcessed(string jsonPath)
        {
            var groupFolder = Path.GetDirectoryName(jsonPath);
            if (string.IsNullOrWhiteSpace(groupFolder))
            {
                return;
            }

            var processedFolder = Path.Combine(groupFolder, "processed");
            Directory.CreateDirectory(processedFolder);

            var destinationPath = Path.Combine(processedFolder, Path.GetFileName(jsonPath));
            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }

            File.Move(jsonPath, destinationPath);
            Console.WriteLine($"[Status] Moved JSON to processed: {destinationPath}");
        }

        private static List<AccessExportRow> SaveProjectExcel(
            string outputPath,
            string groupName,
            string projectFileName,
            IReadOnlyList<DeviceExportRow> rows,
            DateTime recordTimestamp,
            SiemensMetadataService metadataService)
        {
            EnsureFreshFile(outputPath);
            var accessRows = new List<AccessExportRow>(rows.Count);

            using (var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                EnsureStyles(workbookPart);

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                var columns = new Columns(
                    new Column { Min = 1, Max = 1, Width = 20, CustomWidth = true }, // Timestamp
                    new Column { Min = 2, Max = 2, Width = 12, CustomWidth = true }, // Group
                    new Column { Min = 3, Max = 3, Width = 40, CustomWidth = true }, // Project
                    new Column { Min = 4, Max = 4, Width = 28, CustomWidth = true }, // Device
                    new Column { Min = 5, Max = 5, Width = 28, CustomWidth = true }, // DeviceItem
                    new Column { Min = 6, Max = 6, Width = 24, CustomWidth = true }, // OrderNumber
                    new Column { Min = 7, Max = 7, Width = 14, CustomWidth = true }, // Firmware
                    new Column { Min = 8, Max = 8, Width = 12, CustomWidth = true }, // Image
                    new Column { Min = 9, Max = 9, Width = 20, CustomWidth = true }, // Lifecycle
                    new Column { Min = 10, Max = 10, Width = 60, CustomWidth = true }, // Description
                    new Column { Min = 11, Max = 11, Width = 50, CustomWidth = true } // URL
                );
                worksheetPart.Worksheet = new Worksheet(columns, sheetData);

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Devices"
                };
                sheets.Append(sheet);

                var headerRow = new Row { CustomHeight = true, Height = 15D };
                headerRow.Append(
                    CreateCell("Timestamp", 1),
                    CreateCell("Group", 1),
                    CreateCell("Project", 1),
                    CreateCell("Device", 1),
                    CreateCell("DeviceItem", 1),
                    CreateCell("OrderNumber", 1),
                    CreateCell("Firmware", 1),
                    CreateCell("Image", 8),
                    CreateCell("Lifecycle", 8),
                    CreateCell("Description", 8),
                    CreateCell("URL", 1));
                sheetData.Append(headerRow);

                var timestampValue = recordTimestamp.ToString("yyyy-MM-dd HH:mm:ss");
                var hyperlinkRequests = new List<HyperlinkRequest>();
                var total = rows.Count;
                var rowNumber = 2;
                const string urlColumn = "K";

                for (var i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    var orderNumber = Normalize(row.OrderNumber);
                    Console.WriteLine();
                    Console.WriteLine($"[Status] Device {i + 1}/{total}: {Normalize(row.DeviceItem)} ({orderNumber})");
                    var metadata = metadataService.Get(orderNumber);
                    var url = BuildSiemensUrl(orderNumber);
                    var lifecycleStyleIndex = GetLifecycleStyleIndex(metadata.Lifecycle);
                    var imageCellValue = string.Empty;

                    var dataRow = new Row { CustomHeight = true, Height = 60D };
                    dataRow.Append(
                        CreateCell(timestampValue, 2),
                        CreateCell(groupName, 2),
                        CreateCell(projectFileName, 2),
                        CreateCell(Normalize(row.Device), 2),
                        CreateCell(Normalize(row.DeviceItem), 2),
                        CreateCell(orderNumber, 2),
                        CreateCell(Normalize(row.Firmware), 2),
                        CreateCell(imageCellValue, 2),
                        CreateCell(metadata.Lifecycle, lifecycleStyleIndex),
                        CreateCell(metadata.Description, 2),
                        CreateCell(url, 2));
                    sheetData.Append(dataRow);
                    accessRows.Add(new AccessExportRow
                    {
                        Timestamp = timestampValue,
                        Group = groupName,
                        Project = projectFileName,
                        Device = Normalize(row.Device),
                        DeviceItem = Normalize(row.DeviceItem),
                        OrderNumber = orderNumber,
                        Firmware = Normalize(row.Firmware),
                        Lifecycle = metadata.Lifecycle,
                        Description = metadata.Description,
                        Url = url
                    });

                    if (!string.IsNullOrWhiteSpace(url))
                    {
                        hyperlinkRequests.Add(new HyperlinkRequest
                        {
                            CellReference = urlColumn + rowNumber,
                            Url = url
                        });
                    }

                    if (!string.IsNullOrWhiteSpace(metadata.Image) && File.Exists(metadata.Image))
                    {
                        AddImageToWorksheet(worksheetPart, metadata.Image, rowNumber, 7);
                    }

                    rowNumber++;
                }

                var lastRow = rowNumber - 1;
                var filterRef = lastRow >= 1 ? $"A1:K{lastRow}" : "A1:K1";
                worksheetPart.Worksheet.Append(new AutoFilter { Reference = filterRef });

                if (hyperlinkRequests.Count > 0)
                {
                    var hyperlinks = new Hyperlinks();
                    foreach (var request in hyperlinkRequests)
                    {
                        Uri uri;
                        if (!Uri.TryCreate(request.Url, UriKind.Absolute, out uri))
                        {
                            continue;
                        }

                        var relation = worksheetPart.AddHyperlinkRelationship(uri, true);
                        hyperlinks.Append(new Hyperlink
                        {
                            Reference = request.CellReference,
                            Id = relation.Id
                        });
                    }

                    if (hyperlinks.ChildElements.Count > 0)
                    {
                        worksheetPart.Worksheet.Append(hyperlinks);
                    }
                }

                EnsureWorksheetDrawingReference(worksheetPart);

                workbookPart.Workbook.Save();
            }

            return accessRows;
        }

        private static void EnsureFreshFile(string path)
        {
            var folder = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(folder))
            {
                Directory.CreateDirectory(folder);
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private static Row CreateRow(params string[] values)
        {
            var row = new Row();

            foreach (var value in values)
            {
                row.Append(new Cell
                {
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(value ?? string.Empty))
                });
            }

            return row;
        }

        private static Cell CreateCell(string value, uint styleIndex)
        {
            return new Cell
            {
                DataType = CellValues.InlineString,
                StyleIndex = styleIndex,
                InlineString = new InlineString(new Text(value ?? string.Empty))
            };
        }

        private static string Normalize(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static string BuildSiemensUrl(string orderNumber)
        {
            if (string.IsNullOrWhiteSpace(orderNumber))
            {
                return string.Empty;
            }

            var compactOrder = new string(orderNumber.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
            return $"https://sieportal.siemens.com/en-ww/products-services/detail/{compactOrder}?tree=CatalogTree";
        }

        private static void EnsureAccessDatabase(string dbPath)
        {
            if (File.Exists(dbPath))
            {
                return;
            }

            var folder = Path.GetDirectoryName(dbPath);
            if (!string.IsNullOrWhiteSpace(folder))
            {
                Directory.CreateDirectory(folder);
            }

            var accessType = Type.GetTypeFromProgID("Access.Application");
            if (accessType == null)
            {
                throw new InvalidOperationException("Access.Application COM class not found.");
            }

            object accessAppCom = null;
            try
            {
                accessAppCom = Activator.CreateInstance(accessType);
                dynamic accessApp = accessAppCom;
                accessApp.Visible = false;
                accessApp.NewCurrentDatabase(dbPath);
            }
            finally
            {
                if (accessAppCom != null)
                {
                    try
                    {
                        ((dynamic)accessAppCom).Quit();
                    }
                    catch
                    {
                    }

                    if (Marshal.IsComObject(accessAppCom))
                    {
                        Marshal.FinalReleaseComObject(accessAppCom);
                    }
                }
            }
        }

        private static void UpsertAccessTable(string dbPath, string tableName, IReadOnlyList<AccessExportRow> rows)
        {
            var accessType = Type.GetTypeFromProgID("Access.Application");
            if (accessType == null)
            {
                throw new InvalidOperationException("Access.Application COM class not found.");
            }

            object accessAppCom = null;
            object databaseCom = null;
            var tableSqlName = WrapAccessIdentifier(tableName);

            try
            {
                accessAppCom = Activator.CreateInstance(accessType);
                dynamic accessApp = accessAppCom;
                accessApp.Visible = false;
                accessApp.OpenCurrentDatabase(dbPath);

                databaseCom = accessApp.CurrentDb();
                dynamic database = databaseCom;

                try
                {
                    database.Execute("DROP TABLE " + tableSqlName);
                }
                catch
                {
                }

                database.Execute(
                    "CREATE TABLE " + tableSqlName + " (" +
                    "[Timestamp] TEXT(50), " +
                    "[GroupName] TEXT(100), " +
                    "[Project] TEXT(255), " +
                    "[Device] TEXT(255), " +
                    "[DeviceItem] TEXT(255), " +
                    "[OrderNumber] TEXT(100), " +
                    "[Firmware] TEXT(100), " +
                    "[Lifecycle] TEXT(255), " +
                    "[Description] LONGTEXT, " +
                    "[URL] LONGTEXT)");

                foreach (var row in rows)
                {
                    var sql =
                        "INSERT INTO " + tableSqlName +
                        " ([Timestamp], [GroupName], [Project], [Device], [DeviceItem], [OrderNumber], [Firmware], [Lifecycle], [Description], [URL]) VALUES (" +
                        "'" + EscapeSqlValue(row.Timestamp) + "', " +
                        "'" + EscapeSqlValue(row.Group) + "', " +
                        "'" + EscapeSqlValue(row.Project) + "', " +
                        "'" + EscapeSqlValue(row.Device) + "', " +
                        "'" + EscapeSqlValue(row.DeviceItem) + "', " +
                        "'" + EscapeSqlValue(row.OrderNumber) + "', " +
                        "'" + EscapeSqlValue(row.Firmware) + "', " +
                        "'" + EscapeSqlValue(row.Lifecycle) + "', " +
                        "'" + EscapeSqlValue(row.Description) + "', " +
                        "'" + EscapeSqlValue(row.Url) + "')";
                    database.Execute(sql);
                }

                database.Close();
            }
            finally
            {
                if (databaseCom != null && Marshal.IsComObject(databaseCom))
                {
                    Marshal.FinalReleaseComObject(databaseCom);
                }

                if (accessAppCom != null)
                {
                    try
                    {
                        ((dynamic)accessAppCom).CloseCurrentDatabase();
                    }
                    catch
                    {
                    }

                    try
                    {
                        ((dynamic)accessAppCom).Quit();
                    }
                    catch
                    {
                    }

                    if (Marshal.IsComObject(accessAppCom))
                    {
                        Marshal.FinalReleaseComObject(accessAppCom);
                    }
                }
            }
        }

        private static string WrapAccessIdentifier(string name)
        {
            return "[" + Normalize(name).Replace("]", "]]") + "]";
        }

        private static string SanitizeAccessTableName(string value)
        {
            var normalized = Normalize(value);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return "Table_Unknown";
            }

            var sanitized = new string(normalized
                .Select(ch => char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_')
                .ToArray())
                .Trim('_');

            if (string.IsNullOrWhiteSpace(sanitized))
            {
                sanitized = "Table_Unknown";
            }

            if (!char.IsLetter(sanitized[0]))
            {
                sanitized = "T_" + sanitized;
            }

            if (sanitized.Length > 64)
            {
                sanitized = sanitized.Substring(0, 64);
            }

            return sanitized;
        }

        private static string EscapeSqlValue(string value)
        {
            return (value ?? string.Empty).Replace("'", "''");
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "unknown";
            }

            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "unknown" : sanitized;
        }

        private static void PrintException(Exception exception)
        {
            var current = exception;
            var depth = 0;

            while (current != null)
            {
                Console.WriteLine($"[{depth}] {current.GetType().FullName}: {current.Message}");
                current = current.InnerException;
                depth++;
            }
        }

        private static void EnsureStyles(WorkbookPart workbookPart)
        {
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(
                    new Font(), // 0 default
                    new Font(new Bold(), new Color { Rgb = HexBinaryValue.FromString("FFFFFF") }) // 1 header
                ),
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }), // 0 required
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), // 1 required
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("4472C4") }) { PatternType = PatternValues.Solid }), // 2 header blue
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("6EE838") }) { PatternType = PatternValues.Solid }), // 3 green
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF33") }) { PatternType = PatternValues.Solid }), // 4 yellow
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("FFC000") }) { PatternType = PatternValues.Solid }), // 5 orange
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("FF7A01") }) { PatternType = PatternValues.Solid }), // 6 orange-red
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("FF0000") }) { PatternType = PatternValues.Solid }), // 7 red
                    new Fill(new PatternFill(new ForegroundColor { Rgb = HexBinaryValue.FromString("00B5E2") }) { PatternType = PatternValues.Solid }) // 8 Siemens cyan
                ),
                new Borders(new Border()),
                new CellFormats(
                    new CellFormat(), // 0 default
                    new CellFormat // 1 header
                    {
                        FontId = 1,
                        FillId = 2,
                        BorderId = 0,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyAlignment = true,
                        Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    },
                    new CellFormat // 2 centered
                    {
                        FontId = 0,
                        FillId = 0,
                        BorderId = 0,
                        ApplyAlignment = true,
                        Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    },
                    new CellFormat { FontId = 0, FillId = 3, BorderId = 0, ApplyFill = true, ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } }, // 3
                    new CellFormat { FontId = 0, FillId = 4, BorderId = 0, ApplyFill = true, ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } }, // 4
                    new CellFormat { FontId = 0, FillId = 5, BorderId = 0, ApplyFill = true, ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } }, // 5
                    new CellFormat { FontId = 0, FillId = 6, BorderId = 0, ApplyFill = true, ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } }, // 6
                    new CellFormat { FontId = 0, FillId = 7, BorderId = 0, ApplyFill = true, ApplyAlignment = true, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true } }, // 7
                    new CellFormat // 8 siemens cyan header
                    {
                        FontId = 1,
                        FillId = 8,
                        BorderId = 0,
                        ApplyFont = true,
                        ApplyFill = true,
                        ApplyAlignment = true,
                        Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                    }
                ));
            stylesPart.Stylesheet.Save();
        }

        private static uint GetLifecycleStyleIndex(string lifecycle)
        {
            var value = Normalize(lifecycle).ToLowerInvariant();
            if (value.Contains("active product")) return 3;
            if (value.Contains("phase out announce")) return 4;
            if (value.Contains("prod. cancellation")) return 5;
            if (value.Contains("prod. discont.")) return 6;
            if (value.Contains("end prod.lifecycl.")) return 7;
            return 2;
        }

        private static void AddImageToWorksheet(WorksheetPart worksheetPart, string imagePath, int rowIndex, int columnIndexZeroBased)
        {
            DrawingsPart drawingsPart;
            Xdr.WorksheetDrawing worksheetDrawing;

            if (worksheetPart.DrawingsPart == null)
            {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetDrawing = new Xdr.WorksheetDrawing();
                drawingsPart.WorksheetDrawing = worksheetDrawing;
            }
            else
            {
                drawingsPart = worksheetPart.DrawingsPart;
                worksheetDrawing = drawingsPart.WorksheetDrawing;
                if (worksheetDrawing == null)
                {
                    worksheetDrawing = new Xdr.WorksheetDrawing();
                    drawingsPart.WorksheetDrawing = worksheetDrawing;
                }
            }

            var imageType = ImagePartType.Jpeg;
            var ext = Path.GetExtension(imagePath).ToLowerInvariant();
            if (ext == ".png") imageType = ImagePartType.Png;
            else if (ext == ".gif") imageType = ImagePartType.Gif;
            else if (ext == ".bmp") imageType = ImagePartType.Bmp;

            var imagePart = drawingsPart.AddImagePart(imageType);
            using (var stream = File.OpenRead(imagePath))
            {
                imagePart.FeedData(stream);
            }

            var imagePartId = drawingsPart.GetIdOfPart(imagePart);
            var pictureId = (uint)(worksheetDrawing.ChildElements.Count + 1);
            const long sizeEmu = 609600L; // ~64px

            var picture = new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = pictureId, Name = "Image " + pictureId },
                    new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                ),
                new Xdr.BlipFill(
                    new A.Blip { Embed = imagePartId, CompressionState = A.BlipCompressionValues.Print },
                    new A.Stretch(new A.FillRectangle())
                ),
                new Xdr.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = sizeEmu, Cy = sizeEmu }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            var anchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId(columnIndexZeroBased.ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((rowIndex - 1).ToString()),
                    new Xdr.RowOffset("0")
                ),
                new Xdr.Extent { Cx = sizeEmu, Cy = sizeEmu },
                picture,
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(anchor);
            worksheetDrawing.Save();
        }

        private static void EnsureWorksheetDrawingReference(WorksheetPart worksheetPart)
        {
            if (worksheetPart.DrawingsPart == null)
            {
                return;
            }

            if (worksheetPart.Worksheet.Elements<Drawing>().Any())
            {
                return;
            }

            worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart) });
        }

        private sealed class SiemensMetadataService : IDisposable
        {
            private readonly string _imagesFolder;

            private IPlaywright _playwright;
            private IBrowser _browser;
            private IBrowserContext _context;

            public SiemensMetadataService(string imagesFolder)
            {
                _imagesFolder = imagesFolder;
            }

            public MetadataValue Get(string orderNumber)
            {
                var value = new MetadataValue();
                if (string.IsNullOrWhiteSpace(orderNumber))
                {
                    return value;
                }

                var safeName = SanitizeFileName(orderNumber);
                var cachedImagePath = Path.Combine(_imagesFolder, safeName + ".jpg");
                var hasCachedImage = File.Exists(cachedImagePath) && new FileInfo(cachedImagePath).Length > 0;

                if (hasCachedImage)
                {
                    value.Image = cachedImagePath;
                    Console.WriteLine($"[Meta] {orderNumber} -> image loaded from cache.");
                }

                IPage page = null;

                try
                {
                    EnsureBrowser();

                    var url = BuildSiemensUrl(orderNumber);
                    Console.WriteLine($"[Meta] {orderNumber} -> loading page...");

                    page = _context.NewPageAsync().GetAwaiter().GetResult();
                    for (var attempt = 1; attempt <= 2; attempt++)
                    {
                        if (attempt == 1)
                        {
                            page.GotoAsync(url, new PageGotoOptions
                            {
                                WaitUntil = WaitUntilState.DOMContentLoaded,
                                Timeout = 30000
                            }).GetAwaiter().GetResult();
                        }
                        else
                        {
                            Console.WriteLine($"[Meta] {orderNumber} -> retrying metadata load...");
                            page.ReloadAsync(new PageReloadOptions
                            {
                                WaitUntil = WaitUntilState.DOMContentLoaded,
                                Timeout = 30000
                            }).GetAwaiter().GetResult();
                        }

                        var acceptButton = page.Locator("button:has-text('Accept All Cookies')");
                        if (acceptButton.CountAsync().GetAwaiter().GetResult() > 0)
                        {
                            acceptButton.First.ClickAsync().GetAwaiter().GetResult();
                        }

                        try
                        {
                            page.WaitForSelectorAsync("div.product-metadata-item__label-wrapper.ng-star-inserted", new PageWaitForSelectorOptions
                            {
                                Timeout = 8000
                            }).GetAwaiter().GetResult();
                        }
                        catch
                        {
                        }

                        try
                        {
                            page.WaitForSelectorAsync("div.product-metadata-description", new PageWaitForSelectorOptions
                            {
                                Timeout = 8000
                            }).GetAwaiter().GetResult();
                        }
                        catch
                        {
                        }

                        var lifecycleLocator = page.Locator("div.product-metadata-item__label-wrapper.ng-star-inserted");
                        if (lifecycleLocator.CountAsync().GetAwaiter().GetResult() > 0)
                        {
                            value.Lifecycle = Normalize(lifecycleLocator.First.InnerTextAsync().GetAwaiter().GetResult());
                        }

                        var descriptionLocator = page.Locator("div.product-metadata-description");
                        if (descriptionLocator.CountAsync().GetAwaiter().GetResult() > 0)
                        {
                            value.Description = Normalize(descriptionLocator.First.InnerTextAsync().GetAwaiter().GetResult());
                        }

                        if (!hasCachedImage && string.IsNullOrWhiteSpace(value.Image))
                        {
                            var imageLocator = page.Locator("picture img.fit--contain:visible");
                            if (imageLocator.CountAsync().GetAwaiter().GetResult() == 0)
                            {
                                imageLocator = page.Locator("picture img.fit--contain");
                            }

                            if (imageLocator.CountAsync().GetAwaiter().GetResult() > 0)
                            {
                                var imageUrl = Normalize(imageLocator.First.GetAttributeAsync("src").GetAwaiter().GetResult());
                                if (!string.IsNullOrWhiteSpace(imageUrl))
                                {
                                    Console.WriteLine($"[Meta] {orderNumber} -> downloading image...");
                                    var imageResponse = page.APIRequest.GetAsync(imageUrl, new APIRequestContextOptions
                                    {
                                        Timeout = 30000
                                    }).GetAwaiter().GetResult();

                                    if (imageResponse.Ok)
                                    {
                                        var bytes = imageResponse.BodyAsync().GetAwaiter().GetResult();
                                        if (bytes != null && bytes.Length > 0)
                                        {
                                            Directory.CreateDirectory(_imagesFolder);
                                            File.WriteAllBytes(cachedImagePath, bytes);
                                            value.Image = cachedImagePath;
                                            Console.WriteLine($"[Meta] {orderNumber} -> image saved: {cachedImagePath}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[Meta] {orderNumber} -> image not found on page.");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[Meta] {orderNumber} -> image not found on page.");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"[Meta] {orderNumber} -> image not found on page.");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"[Meta] {orderNumber} -> image not found on page.");
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(value.Lifecycle) || !string.IsNullOrWhiteSpace(value.Description))
                        {
                            break;
                        }
                    }

                    Console.WriteLine($"[Meta] {orderNumber} -> lifecycle: {(string.IsNullOrWhiteSpace(value.Lifecycle) ? "empty" : "ok")}; description: {(string.IsNullOrWhiteSpace(value.Description) ? "empty" : "ok")}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Meta] {orderNumber} -> metadata download failed: {ex.Message}");
                }
                finally
                {
                    if (page != null)
                    {
                        page.CloseAsync().GetAwaiter().GetResult();
                    }
                }

                return value;
            }

            private void EnsureBrowser()
            {
                if (_context != null)
                {
                    return;
                }

                _playwright = Playwright.CreateAsync().GetAwaiter().GetResult();
                _browser = _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                {
                    Headless = false,
                    Channel = "chrome",
                    Args = new[]
                    {
                        "--window-size=360,240",
                        "--window-position=0,0"
                    }
                }).GetAwaiter().GetResult();

                _context = _browser.NewContextAsync(new BrowserNewContextOptions
                {
                    Locale = "en-US",
                    ViewportSize = new ViewportSize
                    {
                        Width = 340,
                        Height = 200
                    }
                }).GetAwaiter().GetResult();
            }

            public void Dispose()
            {
                if (_context != null)
                {
                    _context.CloseAsync().GetAwaiter().GetResult();
                    _context = null;
                }

                if (_browser != null)
                {
                    _browser.CloseAsync().GetAwaiter().GetResult();
                    _browser = null;
                }

                if (_playwright != null)
                {
                    _playwright.Dispose();
                    _playwright = null;
                }
            }
        }

        private sealed class MetadataValue
        {
            public string Image { get; set; }
            public string Lifecycle { get; set; }
            public string Description { get; set; }
        }

        private sealed class AccessExportRow
        {
            public string Timestamp { get; set; }
            public string Group { get; set; }
            public string Project { get; set; }
            public string Device { get; set; }
            public string DeviceItem { get; set; }
            public string OrderNumber { get; set; }
            public string Firmware { get; set; }
            public string Lifecycle { get; set; }
            public string Description { get; set; }
            public string Url { get; set; }
        }

        private class GroupInfo
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public List<string> JsonPaths { get; set; }
            public int JsonCount { get; set; }
        }

        private sealed class HyperlinkRequest
        {
            public string CellReference { get; set; }
            public string Url { get; set; }
        }

        [DataContract]
        private class DeviceExportRow
        {
            [DataMember]
            public string Device { get; set; }

            [DataMember]
            public string DeviceItem { get; set; }

            [DataMember]
            public string OrderNumber { get; set; }

            [DataMember]
            public string Firmware { get; set; }
        }
    }
}
