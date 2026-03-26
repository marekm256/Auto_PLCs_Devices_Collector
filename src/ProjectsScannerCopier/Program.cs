using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProjectsScannerCopier
{
    internal static class Program
    {
        private static readonly Regex MachineCodeRegex = new Regex("[A-Z0-9]{7}", RegexOptions.CultureInvariant);

        private static readonly (string Filter, string TypeName)[] ProjectTypes =
        {
            ("*.s7p", "step7"),
            ("*.ap10", "tia10"),
            ("*.ap11", "tia11"),
            ("*.ap12", "tia12"),
            ("*.ap13", "tia13"),
            ("*.ap14", "tia14"),
            ("*.ap15", "tia15"),
            ("*.ap15_1", "tia15_1"),
            ("*.ap16", "tia16"),
            ("*.ap17", "tia17"),
            ("*.ap18", "tia18"),
            ("*.ap19", "tia19"),
            ("*.ap20", "tia20"),
            ("*.acd", "rockwell")
        };

        private static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("  ProjectsScannerCopier.exe \"C:\\source_root\" \"C:\\target_root\" [\"IgnoreA;IgnoreB\"]");
                return 1;
            }

            var rootPath = args[0];
            var targetDir = args[1];
            var ignoredFolders = ParseIgnoredFolders(args.Length >= 3 ? args[2] : string.Empty);

            if (!Directory.Exists(rootPath))
            {
                Console.WriteLine($"Source path does not exist: {rootPath}");
                return 1;
            }

            var rootCreated = false;
            var summary = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            Console.WriteLine($"Scanning projects in: {rootPath}");
            Console.WriteLine("----------------------------------------");

            foreach (var projectType in ProjectTypes)
            {
                var typeTarget = Path.Combine(targetDir, projectType.TypeName);

                Console.WriteLine();
                Console.WriteLine($"Scanning type: {projectType.Filter}");

                var files = EnumerateFilesSafe(rootPath, projectType.Filter, ignoredFolders).ToList();
                var total = files.Count;
                summary[projectType.TypeName] = total;

                if (total == 0)
                {
                    Console.WriteLine("Nothing found.");
                    continue;
                }

                if (!rootCreated)
                {
                    Directory.CreateDirectory(targetDir);
                    rootCreated = true;
                }

                if (!Directory.Exists(typeTarget))
                {
                    Directory.CreateDirectory(typeTarget);
                }

                var copied = 0;
                for (var index = 0; index < files.Count; index++)
                {
                    var filePath = files[index];
                    var sourceFolder = Path.GetDirectoryName(filePath) ?? string.Empty;
                    var machineName = ResolveMachineName(sourceFolder);
                    var destination = BuildVersionedDestination(typeTarget, machineName);

                    Console.WriteLine($"[{index + 1}/{total}] Copying: {sourceFolder} -> {destination}");

                    try
                    {
                        CopyDirectory(sourceFolder, destination);
                        copied++;
                    }
                    catch (Exception copyException)
                    {
                        Console.WriteLine($"  Copy failed: {copyException.Message}");
                    }
                }

                Console.WriteLine($"TOTAL COPIED: {copied}");
            }

            Console.WriteLine();
            Console.WriteLine("========= SUMMARY =========");

            foreach (var key in summary.Keys.OrderBy(key => key, StringComparer.OrdinalIgnoreCase))
            {
                Console.WriteLine($"{key,-10} : {summary[key]}");
            }

            Console.WriteLine("===========================");
            Console.WriteLine("Done.");
            return 0;
        }

        private static HashSet<string> ParseIgnoredFolders(string ignoredFoldersArgument)
        {
            var ignoredFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Source Codes Archive"
            };

            if (string.IsNullOrWhiteSpace(ignoredFoldersArgument))
            {
                return ignoredFolders;
            }

            foreach (var entry in ignoredFoldersArgument.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmed = entry.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed))
                {
                    ignoredFolders.Add(trimmed);
                }
            }

            return ignoredFolders;
        }

        private static IEnumerable<string> EnumerateFilesSafe(string rootPath, string filter, HashSet<string> ignoredFolders)
        {
            var pending = new Stack<string>();
            pending.Push(rootPath);

            while (pending.Count > 0)
            {
                var current = pending.Pop();
                if (ShouldSkipPath(current, ignoredFolders))
                {
                    continue;
                }

                IEnumerable<string> files;
                try
                {
                    files = Directory.EnumerateFiles(current, filter, SearchOption.TopDirectoryOnly);
                }
                catch
                {
                    files = Enumerable.Empty<string>();
                }

                foreach (var file in files)
                {
                    if (!ShouldSkipPath(file, ignoredFolders))
                    {
                        yield return file;
                    }
                }

                IEnumerable<string> directories;
                try
                {
                    directories = Directory.EnumerateDirectories(current, "*", SearchOption.TopDirectoryOnly);
                }
                catch
                {
                    directories = Enumerable.Empty<string>();
                }

                foreach (var directory in directories)
                {
                    pending.Push(directory);
                }
            }
        }

        private static bool ShouldSkipPath(string path, HashSet<string> ignoredFolders)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return true;
            }

            if (path.IndexOf('$') >= 0)
            {
                return true;
            }

            foreach (var folderName in ignoredFolders)
            {
                if (path.IndexOf(folderName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string ResolveMachineName(string sourceFolder)
        {
            var parts = sourceFolder
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .ToArray();

            var seg4 = GetSegment(parts, 4);
            var seg5 = GetSegment(parts, 5);
            var seg6 = GetSegment(parts, 6);

            if (MachineCodeRegex.IsMatch(seg4))
            {
                return seg4;
            }

            if (MachineCodeRegex.IsMatch(seg5))
            {
                return seg5;
            }

            return seg6;
        }

        private static string GetSegment(string[] parts, int index)
        {
            return index >= 0 && index < parts.Length ? parts[index] : string.Empty;
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (var value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            return string.Empty;
        }

        private static string BuildVersionedDestination(string typeTarget, string machineName)
        {
            var baseDestination = Path.Combine(typeTarget, SanitizePathSegment(machineName));
            var destination = baseDestination;
            var version = 2;

            while (Directory.Exists(destination) || File.Exists(destination))
            {
                destination = baseDestination + "_v" + version;
                version++;
            }

            return destination;
        }

        private static string SanitizePathSegment(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "UNKNOWN";
            }

            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitized = new string(value.Select(character => invalidChars.Contains(character) ? '_' : character).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "UNKNOWN" : sanitized;
        }

        private static void CopyDirectory(string sourceDirectory, string destinationDirectory)
        {
            var source = new DirectoryInfo(sourceDirectory);
            if (!source.Exists)
            {
                throw new DirectoryNotFoundException($"Source directory not found: {sourceDirectory}");
            }

            Directory.CreateDirectory(destinationDirectory);

            foreach (var file in source.GetFiles())
            {
                var destinationFilePath = Path.Combine(destinationDirectory, file.Name);
                file.CopyTo(destinationFilePath, true);
            }

            foreach (var subDirectory in source.GetDirectories())
            {
                var destinationSubDirectory = Path.Combine(destinationDirectory, subDirectory.Name);
                CopyDirectory(subDirectory.FullName, destinationSubDirectory);
            }
        }
    }
}
