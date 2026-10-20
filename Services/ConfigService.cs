using System;
using System.IO;
using System.Text.Json;
using HexxcorpDocumentReviewUtility.Models;

namespace HexxcorpDocumentReviewUtility.Services
{
    /// <summary>
    /// Loads and saves the application's local JSON configuration
    /// (<c>config.json</c>) from/to the application's base directory.
    /// All failures are handled gracefully so a missing or corrupt config
    /// file never prevents the application from starting.
    /// </summary>
    public static class ConfigService
    {
        private static readonly string ConfigFilePath =
            Path.Combine(AppContext.BaseDirectory, "config.json");

        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            WriteIndented = true
        };

        /// <summary>
        /// Loads the configuration file if present and valid; otherwise
        /// returns a fresh default configuration.
        /// </summary>
        public static AppConfig Load()
        {
            try
            {
                if (!File.Exists(ConfigFilePath))
                {
                    return new AppConfig();
                }

                string json = File.ReadAllText(ConfigFilePath);
                var config = JsonSerializer.Deserialize<AppConfig>(json, SerializerOptions);
                return config ?? new AppConfig();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                // Config corruption or access issues should never crash the app.
                return new AppConfig();
            }
        }

        /// <summary>
        /// Persists the configuration to disk. Failures are swallowed (and
        /// surfaced via the returned boolean) so the UI can optionally notify
        /// the user without crashing on a locked/read-only file.
        /// </summary>
        public static bool Save(AppConfig config)
        {
            try
            {
                string json = JsonSerializer.Serialize(config, SerializerOptions);
                File.WriteAllText(ConfigFilePath, json);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }
        }
    }
}
