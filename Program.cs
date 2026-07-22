using System;
using Spectre.Console;

namespace ExcelSplitter
{
    class Program
    {
        static int Main(string[] args)
        {
            System.Console.OutputEncoding = System.Text.Encoding.UTF8;
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

                Console.ForegroundColor = ConsoleColor.Green;

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

static void ShowIntro()
    {
        AnsiConsole.Clear();

        // Animated Figlet Logo
        AnsiConsole.Write(
            new FigletText("Excel Splitter")
                .Centered()
                .Color(Color.DeepSkyBlue1));

        AnsiConsole.WriteLine();

        // Nice Panel
        var panel = new Panel(
    @"[green]Automatically splits Excel rows larger than the configured amount.[/]

[yellow]✓[/] Preserves Formatting
[yellow]✓[/] Preserves Merged Cells
[yellow]✓[/] Preserves Row Heights
[yellow]✓[/] Supports XLS and XLSX
[yellow]✓[/] Drag && Drop File Support

[grey]Developer :[/] [lime]Shahab Sadeghi[/]
[grey]Mobile    :[/] [aqua]09177388400[/]")
        {
            Border = BoxBorder.Double,
            Header = new PanelHeader("[bold yellow]Excel Amount Splitter[/]"),
            BorderStyle = new Style(Color.DeepSkyBlue1)
        };

        AnsiConsole.Write(panel);

        AnsiConsole.WriteLine();

        // Loading animation
        AnsiConsole.Status()
            .Spinner(Spinner.Known.Aesthetic)
            .SpinnerStyle(Style.Parse("green"))
            .Start("Loading components...", ctx =>
            {
                Thread.Sleep(1000);
                ctx.Status("Initializing Excel Engine...");
                Thread.Sleep(1100);
                ctx.Status("Preparing...");
                Thread.Sleep(800);
            });

        AnsiConsole.MarkupLine(
            "[bold green]>> you must[/] Drag an Excel file onto this program.");
    }



    private static void WaitForExit()
        {
            Console.WriteLine("\nPress Enter to exit...");
            Console.ReadLine();
        }
    }
}
