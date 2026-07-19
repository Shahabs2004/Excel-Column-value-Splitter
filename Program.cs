using System;

namespace ExcelSplitter
{
    class Program
    {
        static int Main(string[] args)
        {
            ShowIntro();

            string exeDir = AppContext.BaseDirectory;

            using var bootstrapLogger = new Logger(exeDir);
            var config = AppConfig.LoadOrCreate(exeDir, bootstrapLogger);

            // If LogFolder is configured to a different path, a second logger is created there.
            string logFolder = string.IsNullOrWhiteSpace(config.LogFolder) ? exeDir : config.LogFolder;
            Logger logger = logFolder == exeDir ? bootstrapLogger : new Logger(logFolder);

            try
            {
                if (args.Length == 0)
                {
                    logger.Error("No file specified. Drag and drop Excel file(s) (.xls or .xlsx) onto the program, or pass the path as an argument.");
                    WaitForExit();
                    return 1;
                }

                var processor = new ExcelProcessor(config, logger);
                int successCount = 0, failCount = 0;

                foreach (var inputPath in args)
                {
                    logger.Info($"====== Starting to process file: {inputPath} ======");
                    var result = processor.ProcessFile(inputPath);

                    if (result.Success)
                    {
                        logger.Success($"Output file saved: {result.OutputPath} (total rows split: {result.TotalSplitRows})");
                        successCount++;
                    }
                    else
                    {
                        logger.Error($"Processing '{inputPath}' failed: {result.ErrorMessage}");
                        failCount++;
                    }
                }

                logger.Info($"Done. Succeeded: {successCount} | Failed: {failCount}");
                logger.Info($"Full log file: {logger.LogFilePath}");

                WaitForExit();
                return failCount == 0 ? 0 : 2;
            }
            finally
            {
                if (!ReferenceEquals(logger, bootstrapLogger))
                    logger.Dispose();
            }
        }

        private static void ShowIntro()
        {
            var prevColor = Console.ForegroundColor;

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(@"
  _____                _        _                            _     _____       _ _ _   _            
 | ____|_  _____ ___| |      / \   _ __ ___   ___  _   _ _ __ | |_  / ___| _ __ | (_) |_| |_ ___ _ __ 
 |  _| \ \/ / __/ _ \ |     / _ \ | '_ ` _ \ / _ \| | | | '_ \| __| \___ \| '_ \| | | __| __/ _ \ '__|
 | |___ >  < (_|  __/ |    / ___ \| | | | | | (_) | |_| | | | | |_   ___) | |_) | | | |_| ||  __/ |   
 |_____/_/\_\___\___|_|   /_/   \_\_| |_| |_|\___/ \__,_|_| |_|\__| |____/| .__/|_|_|\__|\__\___|_|   
                                                                           |_|                          ");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("  " + new string('=', 100));

            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("   Excel Amount Splitter");
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("   Automatically splits any payment row above a configured limit into multiple");
            Console.WriteLine("   rows (each within the limit), while preserving the original file's formatting,");
            Console.WriteLine("   merged cells, and row heights. Supports both legacy .xls and modern .xlsx files.");

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("  " + new string('-', 100));

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("   Programmer: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Shahab Sadeghi");

            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("  " + new string('=', 100));
            Console.WriteLine();

            Console.ForegroundColor = prevColor;
        }

        private static void WaitForExit()
        {
            Console.WriteLine("\nPress Enter to exit...");
            Console.ReadLine();
        }
    }
}
