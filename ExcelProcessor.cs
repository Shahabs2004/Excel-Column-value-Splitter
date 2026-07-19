using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelSplitter
{
    /// <summary>
    /// Result of processing a single file, used for reporting in Program.cs.
    /// </summary>
    public class FileProcessResult
    {
        public bool Success { get; set; }
        public string? OutputPath { get; set; }
        public int TotalSplitRows { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Opens a single Excel file (xls or xlsx), splits over-limit amount rows in every sheet,
    /// and saves the result to a new output file. Uses NPOI so both formats are supported.
    /// </summary>
    public class ExcelProcessor
    {
        private const int XlsMaxRows = 65536; // row limit of the legacy .xls (BIFF8) format

        private readonly AppConfig _config;
        private readonly Logger _logger;

        public ExcelProcessor(AppConfig config, Logger logger)
        {
            _config = config;
            _logger = logger;
        }

        public FileProcessResult ProcessFile(string inputPath)
        {
            if (!File.Exists(inputPath))
            {
                return new FileProcessResult { Success = false, ErrorMessage = $"File not found: {inputPath}" };
            }

            bool isXls = Path.GetExtension(inputPath).Equals(".xls", StringComparison.OrdinalIgnoreCase);

            // The whole file is read into memory first so the workbook doesn't depend on a disk
            // stream while it's being processed (important for xlsx, which NPOI reads as an OPC package).
            byte[] fileBytes = File.ReadAllBytes(inputPath);
            using var memoryStream = new MemoryStream(fileBytes);

            IWorkbook workbook;
            try
            {
                workbook = WorkbookFactory.Create(memoryStream);
            }
            catch (Exception ex)
            {
                return new FileProcessResult { Success = false, ErrorMessage = $"Could not open the file (invalid format?): {ex.Message}" };
            }

            try
            {
                int totalSplit = 0;
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    var sheet = workbook.GetSheetAt(s);
                    totalSplit += ProcessSheet(sheet);
                }

                if (isXls && ExceedsXlsRowLimit(workbook))
                {
                    return new FileProcessResult
                    {
                        Success = false,
                        ErrorMessage = $"The resulting row count exceeds the legacy .xls limit ({XlsMaxRows:N0} rows). " +
                                       "Please convert the input file to .xlsx first and run again."
                    };
                }

                string outputPath = BuildOutputPath(inputPath, _config.OutputFolder);
                using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(outFs);
                }

                return new FileProcessResult { Success = true, OutputPath = outputPath, TotalSplitRows = totalSplit };
            }
            finally
            {
                workbook.Close();
            }
        }

        private static bool ExceedsXlsRowLimit(IWorkbook workbook)
        {
            for (int s = 0; s < workbook.NumberOfSheets; s++)
            {
                if (workbook.GetSheetAt(s).LastRowNum + 1 > XlsMaxRows)
                    return true;
            }
            return false;
        }

        private int ProcessSheet(ISheet sheet)
        {
            var headerRow = sheet.GetRow(0);
            if (headerRow == null)
            {
                _logger.Warning($"Sheet '{sheet.SheetName}' has no header row; skipped.");
                return 0;
            }

            var cols = _config.Columns;

            VerifyHeader(headerRow, cols.AmountColumnIndex, _config.AmountHeaderHints, "Amount");
            VerifyHeader(headerRow, cols.AccountColumnIndex, _config.AccountHeaderHints, "Account");
            VerifyHeader(headerRow, cols.NameColumnIndex, _config.NameHeaderHints, "Name");

            int firstDataRow = 1;
            int lastRowNum = sheet.LastRowNum; // 0-based

            if (firstDataRow > lastRowNum)
            {
                _logger.Warning($"Sheet '{sheet.SheetName}' has no data rows; skipped.");
                return 0;
            }

            int numCols = Math.Max(headerRow.LastCellNum, (short)(cols.IdColumnIndex + 1));
            int originalDataRowCount = lastRowNum - firstDataRow + 1;
            var progress = new ProgressBar(originalDataRowCount);

            int currentLastRow = lastRowNum;
            int splitCount = 0;
            int footerSkipped = 0;
            int processedCounter = 0;

            // Walk from the bottom row upward; inserting new rows only shifts rows *below*
            // the current row, so rows above (not yet processed) are never disturbed.
            for (int r = lastRowNum; r >= firstDataRow; r--)
            {
                processedCounter++;
                progress.Report(processedCounter, $"Sheet: {sheet.SheetName}");

                var row = sheet.GetRow(r);
                if (row == null) continue;

                var amountCell = row.GetCell(cols.AmountColumnIndex);
                var accountCell = row.GetCell(cols.AccountColumnIndex);
                var nameCell = row.GetCell(cols.NameColumnIndex);

                if (!CellHelper.TryGetDecimal(amountCell, out decimal amount))
                    continue; // completely empty row, or no valid numeric amount

                bool isDataRow = !CellHelper.IsBlank(accountCell) || !CellHelper.IsBlank(nameCell);

                if (!isDataRow)
                {
                    // Has an amount but both account and name are blank => likely a total/footer row; do not split it.
                    footerSkipped++;
                    _logger.Info($"Row {r + 1} detected as a total/footer row (amount={amount:N0}) and was not split.");
                    continue;
                }

                if (amount <= _config.MaxAmount)
                    continue; // no need to split

                var chunks = SplitAmount(amount, _config.MaxAmount);
                int extra = chunks.Count - 1;
                if (extra <= 0) continue;

                InsertSplitRows(sheet, row, r, ref currentLastRow, cols.AmountColumnIndex, chunks, numCols);

                splitCount++;
                _logger.Info($"Row {r + 1} (sheet '{sheet.SheetName}'): amount {amount:N0} split into {chunks.Count} rows (max {_config.MaxAmount:N0} each).");
            }

            progress.Report(originalDataRowCount, $"Sheet: {sheet.SheetName} - done");
            _logger.Info($"Sheet '{sheet.SheetName}': {splitCount} rows split, {footerSkipped} total/footer rows skipped.");
            return splitCount;
        }

        /// <summary>
        /// Inserts new rows directly below the original row using ShiftRows (which automatically
        /// moves merged cells, filters, and named ranges that fall inside the shifted range),
        /// then copies the original row's values/styles into the new rows.
        /// </summary>
        private static void InsertSplitRows(
            ISheet sheet, IRow originalRow, int originalRowIndex,
            ref int currentLastRow, int amountColIndex, List<decimal> chunks, int numCols)
        {
            int extra = chunks.Count - 1;

            if (currentLastRow > originalRowIndex)
                sheet.ShiftRows(originalRowIndex + 1, currentLastRow, extra);

            currentLastRow += extra;

            for (int i = 1; i <= extra; i++)
            {
                int newRowIdx = originalRowIndex + i;
                var newRow = sheet.GetRow(newRowIdx) ?? sheet.CreateRow(newRowIdx);
                CopyRow(originalRow, newRow, numCols);

                var newAmountCell = newRow.GetCell(amountColIndex) ?? newRow.CreateCell(amountColIndex);
                newAmountCell.SetCellValue((double)chunks[i]);
            }

            var origAmountCell = originalRow.GetCell(amountColIndex) ?? originalRow.CreateCell(amountColIndex);
            origAmountCell.SetCellValue((double)chunks[0]);
        }

        private static void CopyRow(IRow src, IRow dst, int numCols)
        {
            dst.Height = src.Height;
            for (int c = 0; c < numCols; c++)
            {
                var srcCell = src.GetCell(c);
                if (srcCell == null) continue;
                var dstCell = dst.GetCell(c) ?? dst.CreateCell(c);
                dstCell.CellStyle = srcCell.CellStyle;
                CellHelper.CopyCellValue(srcCell, dstCell);
            }
        }

        private static List<decimal> SplitAmount(decimal amount, decimal maxAmount)
        {
            var chunks = new List<decimal>();
            decimal remaining = amount;
            while (remaining > 0)
            {
                decimal chunk = Math.Min(remaining, maxAmount);
                chunks.Add(chunk);
                remaining -= chunk;
            }
            return chunks;
        }

        private void VerifyHeader(IRow headerRow, int colIndex, List<string> hints, string fieldLabel)
        {
            var cell = headerRow.GetCell(colIndex);
            var text = CellHelper.GetHeaderText(cell);
            if (!CellHelper.HeaderMatchesAny(text, hints))
            {
                _logger.Warning(
                    $"Column {colIndex + 1} header was expected to be '{fieldLabel}' but its text is '{text}'. " +
                    "Processing continues using the configured fixed column index; please verify the file layout.");
            }
        }

        private static string BuildOutputPath(string inputPath, string configuredOutputFolder)
        {
            string folder = string.IsNullOrWhiteSpace(configuredOutputFolder)
                ? Path.GetDirectoryName(Path.GetFullPath(inputPath))!
                : configuredOutputFolder;

            Directory.CreateDirectory(folder);

            string nameNoExt = Path.GetFileNameWithoutExtension(inputPath);
            string ext = Path.GetExtension(inputPath);
            string candidatePath = Path.Combine(folder, $"{nameNoExt}_output{ext}");

            if (!File.Exists(candidatePath))
                return candidatePath;

            string stamped = $"{nameNoExt}_output_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";
            return Path.Combine(folder, stamped);
        }
    }
}
