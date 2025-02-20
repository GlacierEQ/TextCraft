using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace TextForge
{
    public static class LanguageManager
    {
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        private static readonly string _languageBaseUrl = "https://your-server.com/languages/";
        private static readonly string _localLanguagePath = "Localization/";

        public static async Task DownloadLanguageAsync(string languageCode, IProgress<int> progress = null, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(languageCode))
                throw new ArgumentException("Language code cannot be null or whitespace", nameof(languageCode));

            try
            {
                // Ensure localization directory exists
                Directory.CreateDirectory(_localLanguagePath);

                // Download the language file with streaming
                var response = await _httpClient.GetAsync(
                    $"{_languageBaseUrl}AboutBox.{languageCode}.resx",
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken);
                response.EnsureSuccessStatusCode();

                var filePath = Path.Combine(_localLanguagePath, $"AboutBox.{languageCode}.resx");
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var contentLength = response.Content.Headers.ContentLength ?? 0;
                    var buffer = new byte[8192];
                    long totalRead = 0;
                    using (var stream = await response.Content.ReadAsStreamAsync(cancellationToken))
                    {
                        int read;
                        while ((read = await stream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken)) > 0)
                        {
                            await fs.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
                            totalRead += read;
                            if (contentLength > 0)
                            {
                                progress?.Report((int)((double)totalRead / contentLength * 100));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
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
                foreach (var file in Directory.GetFiles(_localLanguagePath, "AboutBox.*.resx"))
                {
                    var langCode = Path.GetFileNameWithoutExtension(file).Split('.')[1];
                    if (!languages.Contains(langCode))
                        languages.Add(langCode);
                }
            }
            return languages;
        }

        public static bool IsLanguageDownloaded(string languageCode)
        {
            var filePath = Path.Combine(_localLanguagePath, $"AboutBox.{languageCode}.resx");
            return File.Exists(filePath);
        }
    }
}
