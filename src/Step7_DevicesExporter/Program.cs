using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace Step7_DevicesExporter
{
    internal static class Program
    {
        private static readonly Regex MlfbPattern =
            new Regex(@"[A-Z0-9]{4} [A-Z0-9]{3}-[A-Z0-9]{5}-[A-Z0-9]{4}", RegexOptions.IgnoreCase);
        private static readonly Regex ReIos = new Regex(
            @"^\s*IOSUBSYSTEM\s+(\d+)(?:\s*,\s*IOADDRESS\s+(\d+))?(?:\s*,\s*SLOT\s+(\d+))?(?:\s*,\s*SUBSLOT\s+(\d+))?(?=,|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex ReDp = new Regex(
            @"^\s*DPSUBSYSTEM\s+(\d+)(?:\s*,\s*DPADDRESS\s+(\d+))?(?:\s*,\s*SLOT\s+(\d+))?(?:\s*,\s*SUBSLOT\s+(\d+))?(?=,|$)",
            RegexOptions.IgnoreCase);
        private static readonly Regex ReRack = new Regex(
            @"^\s*RACK\s+(\d+)(?:\s*,\s*SLOT\s+(\d+))?(?:\s*,\s*SUBSLOT\s+(\d+))?(?=,|$)",
            RegexOptions.IgnoreCase);

        private static int Main(string[] args)
        {
            if (args.Length < 1 || string.IsNullOrWhiteSpace(args[0]))
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("  Step7_DevicesExporter.exe \"C:\\path\\to\\collected_configs\"");
                return 1;
            }

            var inputFolder = args[0];
            var step7Folder = Path.Combine(inputFolder, "step7");

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"[Status] Input folder does not exist: {inputFolder}");
                return 1;
            }

            Console.WriteLine($"[Status] Input folder: {inputFolder}");
            Console.WriteLine($"[Status] Looking for STEP7 folder: {step7Folder}");

            if (!Directory.Exists(step7Folder))
            {
                Console.WriteLine($"[Status] STEP7 folder not found: {step7Folder}");
                return 1;
            }

            var cfgPaths = FindStep7ConfigFiles(step7Folder);
            Console.WriteLine($"[Status] Config files found: {cfgPaths.Count}");

            if (cfgPaths.Count == 0)
            {
                Console.WriteLine("[Status] Nothing to process.");
                Console.WriteLine("[Summary] Successful projects: 0; Projects with errors: 0; Skipped configs: 0");
                return 0;
            }

            var outputFolder = GetOutputFolder(inputFolder);
            Directory.CreateDirectory(outputFolder);
            Console.WriteLine($"[Status] Export folder: {outputFolder}");

            var groups = BuildConfigGroups(cfgPaths, out var skippedConfigs);
            Console.WriteLine($"[Status] Project groups found: {groups.Count}");

            if (groups.Count == 0)
            {
                Console.WriteLine("[Status] No valid config file names to process.");
                Console.WriteLine($"[Summary] Successful projects: 0; Projects with errors: 0; Skipped configs: {skippedConfigs}");
                return 0;
            }

            var successfulProjects = 0;
            var failedProjects = 0;

            for (var index = 0; index < groups.Count; index++)
            {
                var group = groups[index];
                PrintProjectHeader(index + 1, groups.Count, group.ProjectFolderName);
                Console.WriteLine($"[Step] Config files in group: {group.Configs.Count}");

                try
                {
                    var mergedRows = new List<DeviceItemInfo>();

                    foreach (var config in group.Configs)
                    {
                        Console.WriteLine($"[Step] Scanning config: {Path.GetFileName(config.Path)}");
                        var rows = FindDevicesWithDeviceItems(config.Path, config.ProjectName);
                        Console.WriteLine($"[Result] DeviceItems found in config: {rows.Count}");
                        mergedRows.AddRange(rows);
                    }

                    var outputPath = BuildProjectOutputPath(outputFolder, group.ProjectFolderName);
                    var savedCount = SaveProjectDeviceItemsJson(outputPath, mergedRows);

                    Console.WriteLine($"[Result] DeviceItems merged: {mergedRows.Count}");
                    Console.WriteLine($"[Result] DeviceItems saved: {savedCount}");
                    Console.WriteLine($"[Result] JSON: {outputPath}");

                    successfulProjects++;
                    Console.WriteLine("[Status] Project finished.");
                }
                catch (Exception exception)
                {
                    failedProjects++;
                    Console.WriteLine("[Status] Project failed.");
                    PrintException(exception);
                }
                finally
                {
                    PrintProjectFooter();
                }
            }

            Console.WriteLine("[Status] All projects processed.");
            Console.WriteLine($"[Summary] Successful projects: {successfulProjects}; Projects with errors: {failedProjects}; Skipped configs: {skippedConfigs}");
            return 0;
        }

        private static List<string> FindStep7ConfigFiles(string step7Folder)
        {
            return Directory.EnumerateFiles(step7Folder, "*.cfg", SearchOption.TopDirectoryOnly)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<ProjectGroupInfo> BuildConfigGroups(IReadOnlyList<string> configPaths, out int skippedConfigs)
        {
            var groups = new Dictionary<string, ProjectGroupInfo>(StringComparer.OrdinalIgnoreCase);
            skippedConfigs = 0;

            foreach (var configPath in configPaths)
            {
                if (!TryParseConfigFileName(configPath, out var projectFolderName, out _, out var projectName))
                {
                    skippedConfigs++;
                    Console.WriteLine($"[Status] Skipping config with unexpected file name: {Path.GetFileName(configPath)}");
                    continue;
                }

                var key = projectFolderName;
                if (!groups.TryGetValue(key, out var group))
                {
                    group = new ProjectGroupInfo
                    {
                        ProjectFolderName = projectFolderName
                    };
                    groups.Add(key, group);
                }

                group.Configs.Add(new ConfigInfo
                {
                    Path = configPath,
                    ProjectName = projectName
                });
            }

            return groups.Values
                .OrderBy(group => group.ProjectFolderName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool TryParseConfigFileName(
            string configPath,
            out string projectFolderName,
            out string deviceName,
            out string projectName)
        {
            projectFolderName = string.Empty;
            deviceName = string.Empty;
            projectName = string.Empty;

            var fileName = Path.GetFileNameWithoutExtension(configPath) ?? string.Empty;
            var match = Regex.Match(
                fileName,
                @"^(?<projectFolderName>.+)-(?<deviceName>\d+)-(?<projectName>.+)$",
                RegexOptions.IgnoreCase);

            if (match.Success)
            {
                projectFolderName = Normalize(match.Groups["projectFolderName"].Value);
                deviceName = Normalize(match.Groups["deviceName"].Value);
                projectName = Normalize(match.Groups["projectName"].Value).Replace("-", string.Empty);
            }
            else
            {
                var parts = fileName.Split(new[] { '-' }, StringSplitOptions.None);
                if (parts.Length < 3)
                {
                    return false;
                }

                projectFolderName = Normalize(string.Join("-", parts.Take(parts.Length - 2)));
                deviceName = Normalize(parts[parts.Length - 2]);
                projectName = Normalize(parts[parts.Length - 1]).Replace("-", string.Empty);
            }

            return
                !string.IsNullOrWhiteSpace(projectFolderName) &&
                !string.IsNullOrWhiteSpace(deviceName) &&
                !string.IsNullOrWhiteSpace(projectName);
        }

        private static List<DeviceItemInfo> FindDevicesWithDeviceItems(string configPath, string deviceName)
        {
            var rows = new List<DeviceItemInfo>();
            var seenLocationKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in File.ReadLines(configPath, Encoding.GetEncoding("iso-8859-1")))
            {
                var line = (raw ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                if (!line.StartsWith("RACK", StringComparison.OrdinalIgnoreCase) &&
                    !line.StartsWith("DPSUBSYSTEM", StringComparison.OrdinalIgnoreCase) &&
                    !line.StartsWith("IOSUBSYSTEM", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var mlfbMatch = MlfbPattern.Match(line);
                if (!mlfbMatch.Success)
                {
                    continue;
                }

                var locationKey = BuildLocationKey(line, mlfbMatch.Index);
                if (seenLocationKeys.Contains(locationKey))
                {
                    Console.WriteLine($"\t[SKIPPED] Duplicate location: {locationKey} -> MLFB {mlfbMatch.Value}");
                    continue;
                }

                seenLocationKeys.Add(locationKey);

                var extracted = line.Substring(mlfbMatch.Index);
                var parsed = ParseExtractedColumns(extracted);
                if (string.IsNullOrWhiteSpace(parsed.OrderNumber))
                {
                    continue;
                }

                var normalizedDevice = Normalize(deviceName);
                var normalizedDeviceItem = FirstNonEmpty(Normalize(parsed.DeviceItem), "Unknown");
                rows.Add(new DeviceItemInfo
                {
                    Device = normalizedDevice,
                    DeviceItem = normalizedDeviceItem,
                    Path = normalizedDevice + "/" + normalizedDeviceItem,
                    OrderNumber = Normalize(parsed.OrderNumber),
                    Firmware = NormalizeFirmware(parsed.Firmware)
                });
            }

            return rows;
        }

        private static string BuildLocationKey(string line, int mlfbStartIndex)
        {
            var source = line ?? string.Empty;

            var iosMatch = ReIos.Match(source);
            if (iosMatch.Success)
            {
                var ioAddress = iosMatch.Groups[2].Success ? iosMatch.Groups[2].Value : string.Empty;
                var slot = iosMatch.Groups[3].Success ? iosMatch.Groups[3].Value : "0";
                var subslot = iosMatch.Groups[4].Success ? iosMatch.Groups[4].Value : "0";
                return $"IOSUBSYSTEM|{iosMatch.Groups[1].Value}|{ioAddress}|{slot}|{subslot}";
            }

            var dpMatch = ReDp.Match(source);
            if (dpMatch.Success)
            {
                var dpAddress = dpMatch.Groups[2].Success ? dpMatch.Groups[2].Value : string.Empty;
                var slot = dpMatch.Groups[3].Success ? dpMatch.Groups[3].Value : "0";
                var subslot = dpMatch.Groups[4].Success ? dpMatch.Groups[4].Value : "0";
                return $"DPSUBSYSTEM|{dpMatch.Groups[1].Value}|{dpAddress}|{slot}|{subslot}";
            }

            var rackMatch = ReRack.Match(source);
            if (rackMatch.Success)
            {
                var slot = rackMatch.Groups[2].Success ? rackMatch.Groups[2].Value : "0";
                var subslot = rackMatch.Groups[3].Success ? rackMatch.Groups[3].Value : "0";
                return $"RACK|{rackMatch.Groups[1].Value}|{slot}|{subslot}";
            }

            if (mlfbStartIndex > 0 && mlfbStartIndex <= source.Length)
            {
                var prefix = Normalize(source.Substring(0, mlfbStartIndex)).ToLowerInvariant();
                return "HEADER|" + prefix;
            }

            return "HEADER|" + Normalize(source).ToLowerInvariant();
        }

        private static ParsedDeviceItem ParseExtractedColumns(string extracted)
        {
            var columns = extracted
                .Split(new[] { ',' }, StringSplitOptions.None)
                .Select(column => Normalize(column.Trim('"')))
                .ToList();

            if (columns.Count >= 2)
            {
                var tmp = columns[0];
                columns[0] = columns[1];
                columns[1] = tmp;

                var splitSecond = Regex.Split(columns[1], "\"\\s+\"");
                var rebuilt = new List<string> { columns[0] };
                rebuilt.AddRange(splitSecond);
                if (columns.Count > 2)
                {
                    rebuilt.AddRange(columns.Skip(2));
                }

                columns = rebuilt
                    .Select(value => Normalize(value.Trim('"')))
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .ToList();
            }

            var deviceItem = columns.Count > 0 ? columns[0] : string.Empty;
            var orderNumber = columns.Count > 1 ? columns[1] : string.Empty;
            var firmware = columns.Count > 2 ? columns[2] : string.Empty;

            if (string.IsNullOrWhiteSpace(orderNumber))
            {
                return new ParsedDeviceItem();
            }

            var mlfbMatch = MlfbPattern.Match(orderNumber);
            if (!mlfbMatch.Success)
            {
                return new ParsedDeviceItem();
            }

            return new ParsedDeviceItem
            {
                DeviceItem = Normalize(deviceItem),
                OrderNumber = Normalize(mlfbMatch.Value),
                Firmware = Normalize(firmware)
            };
        }

        private static string GetOutputFolder(string inputFolder)
        {
            var inputFullPath = Path.GetFullPath(inputFolder);
            var inputName = new DirectoryInfo(inputFullPath).Name;
            var parent = Directory.GetParent(inputFullPath);

            var outputRoot = string.Equals(inputName, "collected_configs", StringComparison.OrdinalIgnoreCase) && parent != null
                ? parent.FullName
                : inputFullPath;

            return Path.Combine(outputRoot, "collected_components", "step7");
        }

        private static string BuildProjectOutputPath(string outputFolder, string projectFolderName)
        {
            var fileName = $"{SanitizeFileName(projectFolderName)}.json";
            return Path.Combine(outputFolder, fileName);
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "unnamed";
            }

            var invalid = Path.GetInvalidFileNameChars();
            var sanitized = new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "unnamed" : sanitized;
        }

        private static int SaveProjectDeviceItemsJson(string outputPath, IEnumerable<DeviceItemInfo> deviceItems)
        {
            var rows = deviceItems
                .OrderBy(row => row.Device, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.DeviceItem, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.OrderNumber, StringComparer.OrdinalIgnoreCase)
                .Select(row => new DeviceExportRow
                {
                    Device = row.Device,
                    DeviceItem = row.DeviceItem,
                    OrderNumber = row.OrderNumber,
                    Firmware = row.Firmware
                })
                .ToList();

            var serializer = new DataContractJsonSerializer(typeof(List<DeviceExportRow>));
            using (var fileStream = File.Create(outputPath))
            {
                using (var jsonWriter = JsonReaderWriterFactory.CreateJsonWriter(fileStream, Encoding.UTF8, ownsStream: false, indent: true))
                {
                    serializer.WriteObject(jsonWriter, rows);
                    jsonWriter.Flush();
                }
            }

            return rows.Count;
        }

        private static void PrintProjectHeader(int projectIndex, int projectCount, string projectFolderName)
        {
            Console.WriteLine();
            Console.WriteLine(new string('=', 88));
            Console.WriteLine($"[Project] {projectIndex}/{projectCount}");
            Console.WriteLine($"[Group] {projectFolderName}");
            Console.WriteLine(new string('-', 88));
        }

        private static void PrintProjectFooter()
        {
            Console.WriteLine(new string('=', 88));
            Console.WriteLine();
        }

        private static void PrintException(Exception exception)
        {
            Console.WriteLine("Processing failed.");
            var current = exception;
            var depth = 0;

            while (current != null)
            {
                Console.WriteLine($"[{depth}] {current.GetType().FullName}: {current.Message}");
                current = current.InnerException;
                depth++;
            }
        }

        private static string Normalize(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static string NormalizeFirmware(string value)
        {
            var firmware = Normalize(value);
            if (string.IsNullOrWhiteSpace(firmware))
            {
                return string.Empty;
            }

            if (firmware.StartsWith("V", StringComparison.OrdinalIgnoreCase))
            {
                firmware = firmware.Substring(1).TrimStart();
            }

            return firmware;
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (var value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private class ConfigInfo
        {
            public string Path { get; set; }
            public string ProjectName { get; set; }
        }

        private class ProjectGroupInfo
        {
            public string ProjectFolderName { get; set; }
            public List<ConfigInfo> Configs { get; } = new List<ConfigInfo>();
        }

        private class ParsedDeviceItem
        {
            public string DeviceItem { get; set; }
            public string OrderNumber { get; set; }
            public string Firmware { get; set; }
        }

        private class DeviceItemInfo
        {
            public string Device { get; set; }
            public string DeviceItem { get; set; }
            public string Path { get; set; }
            public string OrderNumber { get; set; }
            public string Firmware { get; set; }
        }

        [DataContract]
        public class DeviceExportRow
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
