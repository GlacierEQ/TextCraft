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

        [UserScopedSetting]
        [DefaultSettingValue("4096")]
        public int ChunkSize
        {
            get => (int)this[nameof(ChunkSize)];
            set => this[nameof(ChunkSize)] = value;
        }

        [UserScopedSetting]
        [DefaultSettingValue("true")]
        public bool EnableMemoryOptimization
        {
            get => (bool)this[nameof(EnableMemoryOptimization)];
            set => this[nameof(EnableMemoryOptimization)] = value;
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
