using System;

namespace ExcelSplitter
{
    /// <summary>
    /// Simple text progress bar in the console, redrawn on the current line (\r),
    /// with no external package dependency.
    /// </summary>
    public class ProgressBar
    {
        private readonly int _total;
        private readonly int _barWidth;
        private int _lastPercent = -1;

        public ProgressBar(int total, int barWidth = 40)
        {
            _total = Math.Max(total, 1);
            _barWidth = barWidth;
        }

        /// <summary>Reports current progress. current should be between 1 and total.</summary>
        public void Report(int current, string? suffix = null)
        {
            int percent = (int)(100.0 * current / _total);
            if (percent == _lastPercent && current != _total) return;
            _lastPercent = percent;

            int filled = (int)(_barWidth * (current / (double)_total));
            if (filled > _barWidth) filled = _barWidth;
            string bar = new string('#', filled) + new string('-', _barWidth - filled);

            Console.Write($"\r[{bar}] {percent,3}% ({current}/{_total}) {suffix}   ");
            if (current >= _total) Console.WriteLine();
        }
    }
}
