using System;
using System.Collections.Generic;
using System.Globalization;
using NPOI.SS.UserModel;

namespace ExcelSplitter
{
    /// <summary>
    /// Helper functions for reading cell values, detecting blank cells, and copying
    /// values between cells (works for both xls and xlsx since both use NPOI's shared IWorkbook).
    /// </summary>
    public static class CellHelper
    {
        /// <summary>Reads a decimal value from a cell; supports numeric, parseable text, and formula (cached value).</summary>
        public static bool TryGetDecimal(ICell? cell, out decimal value)
        {
            value = 0;
            if (cell == null) return false;

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    value = (decimal)cell.NumericCellValue;
                    return true;

                case CellType.Formula:
                    try
                    {
                        // Only the cached calculated value is read; the formula itself is not copied.
                        value = (decimal)cell.NumericCellValue;
                        return true;
                    }
                    catch (Exception)
                    {
                        return false;
                    }

                case CellType.String:
                    var text = cell.StringCellValue?.Trim().Replace(",", "");
                    if (!string.IsNullOrEmpty(text) &&
                        decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                        return true;
                    return false;

                default:
                    return false;
            }
        }

        /// <summary>Is this cell blank content-wise (no cell, empty text, or whitespace)?</summary>
        public static bool IsBlank(ICell? cell)
        {
            if (cell == null) return true;
            return cell.CellType switch
            {
                CellType.Blank => true,
                CellType.String => string.IsNullOrWhiteSpace(cell.StringCellValue),
                _ => false
            };
        }

        /// <summary>Returns the header text of a cell for comparison against expected keywords.</summary>
        public static string GetHeaderText(ICell? cell)
        {
            if (cell == null) return "";
            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue?.Trim() ?? "",
                CellType.Numeric => cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                _ => ""
            };
        }

        /// <summary>Does the header text approximately match any of the expected keywords?</summary>
        public static bool HeaderMatchesAny(string headerText, List<string> hints)
        {
            if (string.IsNullOrWhiteSpace(headerText)) return false;
            foreach (var hint in hints)
            {
                if (headerText.Contains(hint, StringComparison.OrdinalIgnoreCase) ||
                    hint.Contains(headerText, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>Copies a cell's value (not its formula) into another cell.</summary>
        public static void CopyCellValue(ICell src, ICell dst)
        {
            switch (src.CellType)
            {
                case CellType.String:
                    dst.SetCellValue(src.StringCellValue);
                    break;
                case CellType.Numeric:
                    dst.SetCellValue(src.NumericCellValue);
                    break;
                case CellType.Boolean:
                    dst.SetCellValue(src.BooleanCellValue);
                    break;
                case CellType.Formula:
                    try
                    {
                        dst.SetCellValue(src.NumericCellValue);
                    }
                    catch (Exception)
                    {
                        try { dst.SetCellValue(src.StringCellValue); }
                        catch (Exception) { /* value could not be copied; destination cell stays blank */ }
                    }
                    break;
                default:
                    break; // Blank or Unknown: nothing to do
            }
        }
    }
}
