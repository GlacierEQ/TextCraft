using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TextForge
{
    public static class LanguageManager
    {
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        private static readonly string _languageBaseUrl = "https://your-server.com/languages/";
        private static readonly string _localLanguagePath = "Localization/";
        private static readonly ConcurrentDictionary<string, bool> _downloadedLanguages = new();
        private static readonly object _downloadLock = new();
        // Updated regex with RegexOptions.Compiled for performance
        private static readonly Regex _languageCodeRegex = new Regex(@"^[a-z]{2}(-[A-Z]{2})?$", RegexOptions.Compiled);

        public static async Task DownloadLanguageAsync(string languageCode, IProgress<int> progress = null)
        {
            if (!IsValidLanguageCode(languageCode))
                throw new ArgumentException("Invalid language code format", nameof(languageCode));

            lock (_downloadLock)
            {
                if (_downloadedLanguages.ContainsKey(languageCode))
                    return;
            }

            try
            {
                // Ensure localization directory exists
                Directory.CreateDirectory(_localLanguagePath);

                // Download the language file
                var response = await _httpClient.GetAsync(
                    $"{_languageBaseUrl}AboutBox.{languageCode}.resx",
                    HttpCompletionOption.ResponseHeadersRead);

                response.EnsureSuccessStatusCode();

                // Save the file
                var filePath = Path.Combine(_localLanguagePath, $"AboutBox.{languageCode}.resx");
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var contentLength = response.Content.Headers.ContentLength ?? 0;
                    var buffer = new byte[8192];
                    var totalRead = 0L;

                    using (var stream = await response.Content.ReadAsStreamAsync())
                    {
                        int read;
                        while ((read = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fs.WriteAsync(buffer, 0, read);
                            totalRead += read;
                            
                            if (contentLength > 0)
                            {
                                var percent = (int)((double)totalRead / contentLength * 100);
                                progress?.Report(percent);
                            }
                        }
                    }
                }

                _downloadedLanguages[languageCode] = true;
            }
            catch (Exception ex)
            {
                _downloadedLanguages.TryRemove(languageCode, out _);
                CommonUtils.DisplayError($"Failed to download language {languageCode}", ex);
                throw;
            }
        }

        public static List<string> GetAvailableLanguages()
        {
            var languages = new List<string>
            {
                "en", // English
                "es"  // Spanish
            };

            if (Directory.Exists(_localLanguagePath))
            {
                languages.AddRange(
                    Directory.GetFiles(_localLanguagePath, "AboutBox.*.resx")
                        .Select(file => Path.GetFileNameWithoutExtension(file).Split('.')[1])
                        .Where(IsValidLanguageCode)
                        .Distinct()
                );
            }

            return languages;
        }

        public static bool IsLanguageDownloaded(string languageCode)
        {
            if (!IsValidLanguageCode(languageCode))
                return false;

            var filePath = Path.Combine(_localLanguagePath, $"AboutBox.{languageCode}.resx");
            return File.Exists(filePath);
        }

        public static bool IsValidLanguageCode(string languageCode)
        {
            return !string.IsNullOrWhiteSpace(languageCode) && 
                   _languageCodeRegex.IsMatch(languageCode);
        }

        public static void CleanupOldLanguages()
        {
            try
            {
                if (!Directory.Exists(_localLanguagePath))
                    return;

                var validFiles = GetAvailableLanguages()
                    .Select(lang => Path.Combine(_localLanguagePath, $"AboutBox.{lang}.resx"))
                    .ToHashSet();

                foreach (var file in Directory.GetFiles(_localLanguagePath, "AboutBox.*.resx"))
                {
                    if (!validFiles.Contains(file))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch (Exception ex)
                        {
                            CommonUtils.DisplayError($"Failed to delete old language file: {file}", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to cleanup old language files", ex);
            }
        }
    }
}
