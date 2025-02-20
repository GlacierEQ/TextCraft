using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.Resources;

namespace TextForge
{
    public static class LocalizationHelper
    {
        private static readonly ConcurrentDictionary<string, ResourceManager> _resourceManagers = new();
        private static readonly object _lock = new();

        public static string GetString(string key, CultureInfo culture = null)
        {
            try
            {
                var resourceManager = GetResourceManagerForCurrentLanguage();
                return resourceManager.GetString(key, culture) ?? $"[[{key}]]";
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError($"Failed to get localized string for key: {key}", ex);
                return $"[[{key}]]";
            }
        }

        private static ResourceManager GetResourceManagerForCurrentLanguage()
        {
            var languageCode = AppSettings.Instance.LanguageCode;
            
            return _resourceManagers.GetOrAdd(languageCode, code => 
            {
                try
                {
                    return new ResourceManager(
                        $"TextForge.Resources.Strings.{code}",
                        typeof(LocalizationHelper).Assembly);
                }
                catch (Exception ex)
                {
                    CommonUtils.DisplayError($"Failed to load resources for language: {code}", ex);
                    return new ResourceManager(
                        "TextForge.Resources.Strings.en",
                        typeof(LocalizationHelper).Assembly);
                }
            });
        }

        public static void ClearCache()
        {
            lock (_lock)
            {
                _resourceManagers.Clear();
            }
        }
    }
}
