using System.Configuration;

namespace TextForge
{
    public sealed class AppSettings : ApplicationSettingsBase
    {
        private static readonly AppSettings _instance = new AppSettings();
        
        public static AppSettings Instance => _instance;

        private AppSettings() { }

        [UserScopedSetting]
        [DefaultSettingValue("en")]
        public string LanguageCode
        {
            get => (string)this[nameof(LanguageCode)];
            set => this[nameof(LanguageCode)] = value;
        }

        public void SaveSettings()
        {
            try
            {
                Save();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to save settings", ex);
            }
        }
    }
}
