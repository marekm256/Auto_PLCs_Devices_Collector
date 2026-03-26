using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace Tia15_21_DevicesExporter
{
    internal static class Program
    {
        private const string ProcessedFolderName = "processed";
        private static readonly IReadOnlyDictionary<TiaVersion, VersionSpec> VersionSpecs =
            new Dictionary<TiaVersion, VersionSpec>
            {
                {
                    TiaVersion.Tia15,
                    new VersionSpec(
                        "-tia15",
                        "tia15",
                        "*.ap15",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V15\PublicAPI\V15\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V15\PublicAPI\V15\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia15_1,
                    new VersionSpec(
                        "-tia15_1",
                        "tia15_1",
                        "*.ap15_1",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V15_1\PublicAPI\V15.1\Siemens.Engineering.dll",
                            @"C:\Program Files\Siemens\Automation\Portal V15_1\PublicAPI\V15_1\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V15_1\PublicAPI\V15.1\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V15_1\PublicAPI\V15_1\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia16,
                    new VersionSpec(
                        "-tia16",
                        "tia16",
                        "*.ap16",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V16\PublicAPI\V16\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V16\PublicAPI\V16\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia17,
                    new VersionSpec(
                        "-tia17",
                        "tia17",
                        "*.ap17",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V17\PublicAPI\V17\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V17\PublicAPI\V17\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia18,
                    new VersionSpec(
                        "-tia18",
                        "tia18",
                        "*.ap18",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V18\PublicAPI\V18\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V18\PublicAPI\V18\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia19,
                    new VersionSpec(
                        "-tia19",
                        "tia19",
                        "*.ap19",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V19\PublicAPI\V19\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V19\PublicAPI\V19\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia20,
                    new VersionSpec(
                        "-tia20",
                        "tia20",
                        "*.ap20",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V20\PublicAPI\V20\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V20\PublicAPI\V20\Siemens.Engineering.dll"
                        })
                },
                {
                    TiaVersion.Tia21,
                    new VersionSpec(
                        "-tia21",
                        "tia21",
                        "*.ap21",
                        new[]
                        {
                            @"C:\Program Files\Siemens\Automation\Portal V21\PublicAPI\V21\Siemens.Engineering.dll",
                            @"C:\Program Files (x86)\Siemens\Automation\Portal V21\PublicAPI\V21\Siemens.Engineering.dll"
                        })
                }
            };

        private static readonly Regex OrderNumberFromTypeIdentifier =
            new Regex(@"^OrderNumber:(?<order>[^/]*)(?:/(?<fw>[^/]*))?", RegexOptions.IgnoreCase);

        private static readonly string[] OrderAttributeNames = { "OrderNumber", "MLFB", "ArticleNumber", "OrderNo" };
        private static readonly string[] FirmwareAttributeNames = { "FirmwareVersion", "Firmware", "Version", "Revision" };

        private static int Main(string[] args)
        {
            if (args.Length < 2 || string.IsNullOrWhiteSpace(args[0]) || string.IsNullOrWhiteSpace(args[1]))
            {
                PrintUsage();
                return 1;
            }

            if (!TryParseVersion(args[0], out var version))
            {
                Console.WriteLine($"[Status] Unsupported version option: {args[0]}");
                PrintUsage();
                return 1;
            }

            var versionSpec = VersionSpecs[version];
            var inputFolder = args[1];
            var versionFolderName = versionSpec.FolderName;
            var projectExtension = versionSpec.ProjectExtension;
            var opennessDllPath = ResolveOpennessDllPath(versionSpec);

            if (string.IsNullOrWhiteSpace(opennessDllPath))
            {
                Console.WriteLine("[Status] Openness DLL not found. Checked:");
                foreach (var candidate in versionSpec.OpennessDllCandidates)
                {
                    Console.WriteLine($"[Status]   {candidate}");
                }
                return 1;
            }

            var tiaFolder = Path.Combine(inputFolder, versionFolderName);
            var processedFolder = Path.Combine(tiaFolder, ProcessedFolderName);

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"[Status] Input folder does not exist: {inputFolder}");
                return 1;
            }

            Console.WriteLine($"[Status] Version: {versionFolderName}");
            Console.WriteLine($"[Status] Input folder: {inputFolder}");
            Console.WriteLine($"[Status] Looking for folder: {tiaFolder}");
            Console.WriteLine($"[Status] Using Openness DLL: {opennessDllPath}");

            if (!Directory.Exists(tiaFolder))
            {
                Console.WriteLine($"[Status] Version folder not found: {tiaFolder}");
                return 1;
            }

            Console.WriteLine($"[Status] Searching project folders in: {tiaFolder}");
            var projectPaths = FindProjectFiles(tiaFolder, projectExtension);
            Console.WriteLine($"[Status] Projects found: {projectPaths.Count}");

            if (projectPaths.Count == 0)
            {
                Console.WriteLine("[Status] Nothing to process.");
                Console.WriteLine("[Summary] Successful projects: 0; Projects with errors: 0");
                return 0;
            }

            var outputFolder = GetOutputFolder(inputFolder, versionFolderName);
            Directory.CreateDirectory(outputFolder);
            Directory.CreateDirectory(processedFolder);
            Console.WriteLine($"[Status] Export folder: {outputFolder}");
            Console.WriteLine($"[Status] Processed folder: {processedFolder}");

            object tiaPortal = null;

            try
            {
                Console.WriteLine("[Status] Opening TIA Portal...");
                var engineeringAssembly = Assembly.LoadFrom(opennessDllPath);
                tiaPortal = CreateTiaPortal(engineeringAssembly);
                Console.WriteLine("[Status] TIA Portal opened.");

                var successfulProjects = 0;
                var failedProjects = 0;

                for (var index = 0; index < projectPaths.Count; index++)
                {
                    var projectPath = projectPaths[index];
                    var projectFolderPath = Path.GetDirectoryName(projectPath);
                    object project = null;
                    var processedSuccessfully = false;

                    PrintProjectHeader(index + 1, projectPaths.Count, projectPath);

                    try
                    {
                        Console.WriteLine("[Step] Opening project...");
                        project = OpenProject(tiaPortal, projectPath);

                        Console.WriteLine("[Step] Scanning devices...");
                        var deviceItems = FindDevicesWithDeviceItems(project);
                        var outputPath = BuildProjectOutputPath(outputFolder, projectPath);
                        var savedCount = SaveProjectDeviceItemsJson(outputPath, deviceItems);

                        Console.WriteLine($"[Result] DeviceItems found: {deviceItems.Count}");
                        Console.WriteLine($"[Result] DeviceItems saved: {savedCount}");
                        Console.WriteLine($"[Result] JSON: {outputPath}");
                        processedSuccessfully = true;
                        successfulProjects++;

                        Console.WriteLine("[Status] Project finished.");
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine("[Status] Project failed.");
                        PrintException(exception);
                        failedProjects++;
                    }
                    finally
                    {
                        if (project != null)
                        {
                            Console.WriteLine("[Step] Closing project...");
                            TryInvokeMethod(project, "Close");
                        }

                        if (processedSuccessfully)
                        {
                            TryMoveProjectFolder(projectFolderPath, processedFolder);
                        }

                        PrintProjectFooter();
                    }
                }

                Console.WriteLine("[Status] All projects processed.");
                Console.WriteLine($"[Summary] Successful projects: {successfulProjects}; Projects with errors: {failedProjects}");
                return 0;
            }
            catch (Exception exception)
            {
                PrintException(exception);
                return 1;
            }
            finally
            {
                if (tiaPortal != null)
                {
                    Console.WriteLine("[Status] Closing TIA Portal...");
                    TryDispose(tiaPortal);
                    Console.WriteLine("[Status] TIA Portal closed.");
                }
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia15 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia15_1 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia16 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia17 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia18 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia19 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia20 \"C:\\path\\to\\projects_root\"");
            Console.WriteLine("  Tia15-21_DevicesExporter.exe -tia21 \"C:\\path\\to\\projects_root\"");
        }

        private static bool TryParseVersion(string value, out TiaVersion version)
        {
            foreach (var entry in VersionSpecs)
            {
                if (string.Equals(value, entry.Value.Option, StringComparison.OrdinalIgnoreCase))
                {
                    version = entry.Key;
                    return true;
                }
            }

            version = default(TiaVersion);
            return false;
        }

        private static string ResolveOpennessDllPath(VersionSpec versionSpec)
        {
            foreach (var candidate in versionSpec.OpennessDllCandidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        private static object CreateTiaPortal(Assembly engineeringAssembly)
        {
            var tiaPortalModeType = engineeringAssembly.GetType("Siemens.Engineering.TiaPortalMode", throwOnError: true);
            var withoutUiMode = Enum.Parse(tiaPortalModeType, "WithoutUserInterface");

            var tiaPortalType = engineeringAssembly.GetType("Siemens.Engineering.TiaPortal", throwOnError: true);
            return Activator.CreateInstance(tiaPortalType, new[] { withoutUiMode });
        }

        private static object OpenProject(object tiaPortal, string projectPath)
        {
            var projects = GetPropertyValue(tiaPortal, "Projects");
            return InvokeMethod(projects, "Open", new FileInfo(projectPath));
        }

        private static List<string> FindProjectFiles(string tiaFolder, string projectExtension)
        {
            var projectFiles = new List<string>();

            foreach (var folder in Directory.EnumerateDirectories(tiaFolder)
                .Where(path => !string.Equals(Path.GetFileName(path), ProcessedFolderName, StringComparison.OrdinalIgnoreCase)))
            {
                var files = Directory.EnumerateFiles(folder, projectExtension, SearchOption.TopDirectoryOnly)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (files.Count == 0)
                {
                    Console.WriteLine($"[Status] Skipping folder without {projectExtension}: {folder}");
                    continue;
                }

                if (files.Count > 1)
                {
                    Console.WriteLine($"[Status] Multiple {projectExtension} found, using first: {folder}");
                }

                projectFiles.Add(files[0]);
            }

            return projectFiles;
        }

        private static List<DeviceItemInfo> FindDevicesWithDeviceItems(object project)
        {
            var deviceItems = new List<DeviceItemInfo>();

            foreach (var device in GetSelectedDevices(project))
            {
                var deviceName = Normalize(AsString(GetPropertyValue(device, "Name")));
                var firstLevelItems = GetPropertyValue(device, "DeviceItems");

                // Intentionally inspect only first-level device items under each device.
                foreach (var item in Enumerate(firstLevelItems))
                {
                    var deviceItemName = Normalize(AsString(GetPropertyValue(item, "Name")));
                    if (!string.IsNullOrEmpty(deviceItemName) &&
                        (deviceItemName.IndexOf("rail", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        deviceItemName.IndexOf("rack", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        continue;
                    }

                    var typeIdentifier = AsString(GetPropertyValue(item, "TypeIdentifier"));
                    var parsed = ParseTypeIdentifier(typeIdentifier);
                    var attributes = ReadAllAttributes(item);
                    var orderNumber = FirstNonEmpty(parsed.OrderNumber, GetAttribute(attributes, OrderAttributeNames));
                    var firmware = NormalizeFirmware(FirstNonEmpty(parsed.Firmware, GetAttribute(attributes, FirmwareAttributeNames)));

                    if (string.IsNullOrWhiteSpace(orderNumber))
                    {
                        continue;
                    }

                    deviceItems.Add(new DeviceItemInfo
                    {
                        Device = deviceName,
                        DeviceItem = deviceItemName,
                        Path = deviceName + "/" + deviceItemName,
                        OrderNumber = orderNumber,
                        Firmware = firmware
                    });
                }
            }

            return deviceItems;
        }

        private static IEnumerable<object> GetSelectedDevices(object project)
        {
            foreach (var device in Enumerate(GetPropertyValue(project, "Devices")))
            {
                yield return device;
            }

            var ungroupedDevicesGroup = GetPropertyValue(project, "UngroupedDevicesGroup");
            if (ungroupedDevicesGroup == null)
            {
                yield break;
            }

            foreach (var device in Enumerate(GetPropertyValue(ungroupedDevicesGroup, "Devices")))
            {
                yield return device;
            }
        }

        private static (string OrderNumber, string Firmware) ParseTypeIdentifier(string typeIdentifier)
        {
            if (string.IsNullOrWhiteSpace(typeIdentifier))
            {
                return (string.Empty, string.Empty);
            }

            var match = OrderNumberFromTypeIdentifier.Match(typeIdentifier);
            if (!match.Success)
            {
                return (string.Empty, string.Empty);
            }

            return (
                match.Groups["order"].Value.Trim(),
                match.Groups["fw"].Value.Trim()
            );
        }

        private static Dictionary<string, string> ReadAllAttributes(object engineeringObject)
        {
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var infos = InvokeMethod(engineeringObject, "GetAttributeInfos");
                foreach (var info in Enumerate(infos))
                {
                    var name = AsString(GetPropertyValue(info, "Name"));
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    try
                    {
                        var raw = InvokeMethod(engineeringObject, "GetAttribute", name);
                        attributes[name] = raw?.ToString() ?? string.Empty;
                    }
                    catch
                    {
                        // Some attributes are not readable in current context; skip.
                    }
                }
            }
            catch
            {
                // Object may not expose dynamic attributes.
            }

            return attributes;
        }

        private static string GetAttribute(Dictionary<string, string> attributes, IEnumerable<string> names)
        {
            foreach (var name in names)
            {
                if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            return string.Empty;
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

        private static string GetOutputFolder(string inputFolder, string versionFolderName)
        {
            var inputFullPath = Path.GetFullPath(inputFolder);
            var inputName = new DirectoryInfo(inputFullPath).Name;
            var parent = Directory.GetParent(inputFullPath);

            var outputRoot = string.Equals(inputName, "collected_projects", StringComparison.OrdinalIgnoreCase) && parent != null
                ? parent.FullName
                : inputFullPath;

            return Path.Combine(outputRoot, "collected_components", versionFolderName);
        }

        private static string BuildProjectOutputPath(string outputFolder, string projectPath)
        {
            var projectFolderName = Path.GetFileName(Path.GetDirectoryName(projectPath));
            var projectFileName = Path.GetFileNameWithoutExtension(projectPath);
            var fileName = $"{SanitizeFileName(projectFolderName)}-{SanitizeFileName(projectFileName)}.json";
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
                .OrderBy(row => row.Path)
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

        private static void TryMoveProjectFolder(string projectFolderPath, string processedFolder)
        {
            if (string.IsNullOrWhiteSpace(projectFolderPath) || !Directory.Exists(projectFolderPath))
            {
                Console.WriteLine("[Status] Project folder move skipped (folder not found).");
                return;
            }

            try
            {
                var destination = BuildUniqueDestinationFolderPath(processedFolder, Path.GetFileName(projectFolderPath));
                Directory.Move(projectFolderPath, destination);
                Console.WriteLine($"[Status] Project folder moved to: {destination}");
            }
            catch (Exception moveException)
            {
                Console.WriteLine($"[Status] Failed to move project folder: {projectFolderPath}");
                Console.WriteLine($"[Status] Move error: {moveException.Message}");
            }
        }

        private static string BuildUniqueDestinationFolderPath(string processedFolder, string folderName)
        {
            var sanitizedFolderName = SanitizeFileName(folderName);
            var candidatePath = Path.Combine(processedFolder, sanitizedFolderName);
            if (!Directory.Exists(candidatePath))
            {
                return candidatePath;
            }

            for (var suffix = 1; suffix < 1000; suffix++)
            {
                candidatePath = Path.Combine(processedFolder, sanitizedFolderName + "_" + suffix);
                if (!Directory.Exists(candidatePath))
                {
                    return candidatePath;
                }
            }

            return Path.Combine(processedFolder, sanitizedFolderName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss"));
        }

        private static void PrintProjectHeader(int projectIndex, int projectCount, string projectPath)
        {
            Console.WriteLine();
            Console.WriteLine(new string('=', 88));
            Console.WriteLine($"[Project] {projectIndex}/{projectCount}");
            Console.WriteLine($"[Path] {projectPath}");
            Console.WriteLine(new string('-', 88));
        }

        private static void PrintProjectFooter()
        {
            Console.WriteLine(new string('=', 88));
            Console.WriteLine();
        }

        private static void PrintException(Exception exception)
        {
            Console.WriteLine("Open project failed.");
            var current = exception;
            var depth = 0;

            while (current != null)
            {
                Console.WriteLine($"[{depth}] {current.GetType().FullName}: {current.Message}");
                current = current.InnerException;
                depth++;
            }
        }


        private static IEnumerable<object> Enumerate(object value)
        {
            if (value is IEnumerable enumerable)
            {
                foreach (var item in enumerable)
                {
                    yield return item;
                }
            }
        }

        private static string AsString(object value)
        {
            return value?.ToString() ?? string.Empty;
        }

        private static object GetPropertyValue(object instance, string propertyName)
        {
            if (instance == null)
            {
                return null;
            }

            var property = instance.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            return property?.GetValue(instance, null);
        }

        private static object InvokeMethod(object instance, string methodName, params object[] arguments)
        {
            if (instance == null)
            {
                throw new InvalidOperationException($"Cannot invoke '{methodName}' on null instance.");
            }

            return instance.GetType().InvokeMember(
                methodName,
                BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null,
                instance,
                arguments);
        }

        private static void TryInvokeMethod(object instance, string methodName, params object[] arguments)
        {
            try
            {
                InvokeMethod(instance, methodName, arguments);
            }
            catch
            {
                // Best effort cleanup.
            }
        }

        private static void TryDispose(object instance)
        {
            if (instance is IDisposable disposable)
            {
                disposable.Dispose();
                return;
            }

            TryInvokeMethod(instance, "Dispose");
        }

        private enum TiaVersion
        {
            Tia15,
            Tia15_1,
            Tia16,
            Tia17,
            Tia18,
            Tia19,
            Tia20,
            Tia21
        }

        private sealed class VersionSpec
        {
            public VersionSpec(string option, string folderName, string projectExtension, IReadOnlyList<string> opennessDllCandidates)
            {
                Option = option;
                FolderName = folderName;
                ProjectExtension = projectExtension;
                OpennessDllCandidates = opennessDllCandidates;
            }

            public string Option { get; }
            public string FolderName { get; }
            public string ProjectExtension { get; }
            public IReadOnlyList<string> OpennessDllCandidates { get; }
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
