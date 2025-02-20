using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace TextForge
{
    public static class LanguageManager
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private static readonly string _languageBaseUrl = "https://your-server.com/languages/";
        private static readonly string _localLanguagePath = "Localization/";

        public static async Task DownloadLanguageAsync(string languageCode)
        {
            try
            {
                // Ensure localization directory exists
                Directory.CreateDirectory(_localLanguagePath);

                // Download the language file
                var response = await _httpClient.GetAsync($"{_languageBaseUrl}AboutBox.{languageCode}.resx");
                response.EnsureSuccessStatusCode();

                // Save the file
                var filePath = Path.Combine(_localLanguagePath, $"AboutBox.{languageCode}.resx");
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fs);
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError($"Failed to download language {languageCode}", ex);
            }
        }

        public static List<string> GetAvailableLanguages()
        {
            var languages = new List<string>();
            
            // Add built-in languages
            languages.Add("en"); // English
            languages.Add("es"); // Spanish

            // Add downloaded languages
            if (Directory.Exists(_localLanguagePath))
            {
                foreach (var file in Directory.GetFiles(_localLanguagePath, "AboutBox.*.resx"))
                {
                    var langCode = Path.GetFileNameWithoutExtension(file).Split('.')[1];
                    if (!languages.Contains(langCode))
                    {
                        languages.Add(langCode);
                    }
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
