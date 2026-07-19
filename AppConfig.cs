using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ExcelSplitter
{
    /// <summary>
    /// Fixed 0-based column mapping matching the real file format:
    /// A=Amount, B=Account (IBAN), C=Name, D=Id.
    /// Since the column order is always fixed, these indices are the primary
    /// source of truth. Header text is only checked as a safety warning.
    /// </summary>
    public class ColumnMapping
    {
        public int AmountColumnIndex { get; set; } = 0;
        public int AccountColumnIndex { get; set; } = 1;
        public int NameColumnIndex { get; set; } = 2;
        public int IdColumnIndex { get; set; } = 3;
    }

    /// <summary>
    /// Editable application settings, loaded from config.json next to the executable.
    /// A default file is created automatically if none exists.
    /// </summary>
    public class AppConfig
    {
        public decimal MaxAmount { get; set; } = 2_000_000_000m;

        public ColumnMapping Columns { get; set; } = new();

        // NOTE: These hint values are the *actual header text found in the source file*
        // (which is in Persian), so they must stay as-is. They are data, not UI text.
        public List<string> AmountHeaderHints { get; set; } = new() { "مبلغ", "Amount" };
        public List<string> AccountHeaderHints { get; set; } = new() { "حساب", "شبا", "IBAN", "Account" };
        public List<string> NameHeaderHints { get; set; } = new() { "نام", "Name" };

        /// <summary>Output folder; if empty, the output is saved next to the input file.</summary>
        public string OutputFolder { get; set; } = "";

        /// <summary>Log folder; if empty, the log is saved next to the executable.</summary>
        public string LogFolder { get; set; } = "";

        private const string ConfigFileName = "config.json";
        private const decimal DefaultMaxAmount = 2_000_000_000m;

        public static AppConfig LoadOrCreate(string baseDirectory, Logger logger)
        {
            string configPath = Path.Combine(baseDirectory, ConfigFileName);
            AppConfig config;

            if (!File.Exists(configPath))
            {
                config = new AppConfig();
                SaveDefault(configPath, config, logger);
                return config;
            }

            try
            {
                string json = File.ReadAllText(configPath);
                config = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                logger.Info($"Settings loaded from '{configPath}'. Max amount: {config.MaxAmount:N0}");
            }
            catch (JsonException ex)
            {
                logger.Error($"Failed to read config.json: {ex.Message}. Using default settings.");
                config = new AppConfig();
            }

            Validate(config, logger);
            return config;
        }

        private static void SaveDefault(string configPath, AppConfig config, Logger logger)
        {
            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configPath, json);
            logger.Info($"config.json not found; a default file was created at '{configPath}'.");
        }

        /// <summary>
        /// Validates critical values. If MaxAmount is zero or negative, the amount-splitting
        /// algorithm would loop forever, so this value must always be positive.
        /// </summary>
        private static void Validate(AppConfig config, Logger logger)
        {
            if (config.MaxAmount <= 0)
            {
                logger.Error($"MaxAmount in config.json is invalid ({config.MaxAmount}). Falling back to default {DefaultMaxAmount:N0}.");
                config.MaxAmount = DefaultMaxAmount;
            }

            var cols = config.Columns;
            var indices = new[] { cols.AmountColumnIndex, cols.AccountColumnIndex, cols.NameColumnIndex, cols.IdColumnIndex };
            foreach (var idx in indices)
            {
                if (idx < 0)
                {
                    logger.Error("One of the column indices in config.json is negative. Falling back to default column mapping.");
                    config.Columns = new ColumnMapping();
                    break;
                }
            }
        }
    }
}
