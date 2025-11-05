using System;
using System.IO;
using System.Linq;

namespace OutlookOSTScanner
{
    class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine("╔" + new string('═', 58) + "╗");
            Console.WriteLine("║ Outlook OST/PST Low-Level File Scanner                  ║");
            Console.WriteLine("║ Deep Structure Validation                                ║");
            Console.WriteLine("╚" + new string('═', 58) + "╝");
            Console.WriteLine();

            // Parse arguments
            var options = ParseArguments(args);

            if (options.ShowHelp)
            {
                ShowHelp();
                return 0;
            }

            // Find files to scan
            string[] filesToScan = options.SpecificFile != null
                ? new[] { options.SpecificFile }
                : FindOSTFiles();

            if (filesToScan.Length == 0)
            {
                Console.WriteLine("No OST/PST files found.");
                return 0;
            }

            Console.WriteLine($"Found {filesToScan.Length} file(s) to scan\n");

            int corruptedCount = 0;
            int scannedCount = 0;

            foreach (string file in filesToScan)
            {
                Console.WriteLine($"\n{"─",60}");
                Console.WriteLine($"Scanning: {Path.GetFileName(file)}");
                Console.WriteLine($"Path: {file}");

                if (File.Exists(file))
                {
                    var fileInfo = new FileInfo(file);
                    Console.WriteLine($"Size: {FormatFileSize(fileInfo.Length)}");
                    Console.WriteLine($"Modified: {fileInfo.LastWriteTime}");
                }

                Console.WriteLine();

                // Perform validation
                var validator = new PSTValidator(file);
                bool isValid = validator.Validate();

                scannedCount++;

                // Display results
                if (!isValid)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("✗ CORRUPTED - File validation failed");
                    Console.ResetColor();
                    corruptedCount++;

                    if (validator.Errors.Count > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("\nErrors:");
                        foreach (var error in validator.Errors)
                        {
                            Console.WriteLine($"  • {error}");
                        }
                        Console.ResetColor();
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("✓ HEALTHY - File validation passed");
                    Console.ResetColor();
                }

                if (validator.Warnings.Count > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\nWarnings:");
                    foreach (var warning in validator.Warnings)
                    {
                        Console.WriteLine($"  • {warning}");
                    }
                    Console.ResetColor();
                }

                // Handle deletion if requested
                if (!isValid && options.DeleteCorrupted)
                {
                    Console.WriteLine();
                    if (DeleteOrBackupFile(file, options.NoBackup))
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine(options.NoBackup
                            ? "✓ File deleted"
                            : "✓ File backed up and renamed");
                        Console.ResetColor();
                    }
                }
                else if (!isValid)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\nUse --delete to remove this corrupted file");
                    Console.ResetColor();
                }
            }

            // Summary
            Console.WriteLine($"\n{"═",60}");
            Console.WriteLine($"Summary:");
            Console.WriteLine($"  Scanned: {scannedCount} file(s)");
            Console.WriteLine($"  Corrupted: {corruptedCount} file(s)");
            Console.WriteLine($"  Healthy: {scannedCount - corruptedCount} file(s)");
            Console.WriteLine($"{"═",60}\n");

            return corruptedCount > 0 ? 1 : 0;
        }

        static string[] FindOSTFiles()
        {
            var files = new System.Collections.Generic.List<string>();

            // Location 1: Local AppData (most common for OST)
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string outlookPath1 = Path.Combine(localAppData, "Microsoft", "Outlook");
            AddFilesFromPath(files, outlookPath1);

            // Location 2: Roaming AppData (sometimes used for PST)
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string outlookPath2 = Path.Combine(appData, "Microsoft", "Outlook");
            AddFilesFromPath(files, outlookPath2);

            // Location 3: Documents folder (user PST files)
            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string outlookPath3 = Path.Combine(documents, "Outlook Files");
            AddFilesFromPath(files, outlookPath3);

            // Remove duplicates
            return files.Distinct().ToArray();
        }

        static void AddFilesFromPath(System.Collections.Generic.List<string> fileList, string path)
        {
            if (Directory.Exists(path))
            {
                try
                {
                    fileList.AddRange(Directory.GetFiles(path, "*.ost", SearchOption.TopDirectoryOnly));
                    fileList.AddRange(Directory.GetFiles(path, "*.pst", SearchOption.TopDirectoryOnly));
                }
                catch (UnauthorizedAccessException)
                {
                    // Skip inaccessible directories
                }
            }
        }

        static bool DeleteOrBackupFile(string filePath, bool noBackup)
        {
            try
            {
                if (noBackup)
                {
                    File.Delete(filePath);
                }
                else
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string backupPath = $"{filePath}.corrupted.{timestamp}.bak";
                    File.Move(filePath, backupPath);
                    Console.WriteLine($"  Backup location: {backupPath}");
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"  Error: {ex.Message}");
                Console.ResetColor();
                return false;
            }
        }

        static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        static (bool ShowHelp, bool DeleteCorrupted, bool NoBackup, string? SpecificFile) ParseArguments(string[] args)
        {
            bool showHelp = args.Contains("--help") || args.Contains("-h");
            bool deleteCorrupted = args.Contains("--delete") || args.Contains("-d");
            bool noBackup = args.Contains("--no-backup");
            string? specificFile = null;

            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "--file" || args[i] == "-f") && i + 1 < args.Length)
                {
                    specificFile = args[i + 1];
                    break;
                }
            }

            return (showHelp, deleteCorrupted, noBackup, specificFile);
        }

        static void ShowHelp()
        {
            Console.WriteLine("Usage: OutlookOSTScanner [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -h, --help          Show this help message");
            Console.WriteLine("  -d, --delete        Delete corrupted files (creates backup by default)");
            Console.WriteLine("  -f, --file <path>   Scan specific file instead of auto-discovery");
            Console.WriteLine("  --no-backup         Permanently delete instead of creating backup");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  OutlookOSTScanner");
            Console.WriteLine("  OutlookOSTScanner --delete");
            Console.WriteLine("  OutlookOSTScanner --file \"C:\\path\\to\\file.ost\"");
            Console.WriteLine("  OutlookOSTScanner --delete --no-backup");
        }
    }
}