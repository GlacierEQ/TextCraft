using System;
using System.Globalization;
using System.Windows.Forms;

namespace TextForge
{
    public partial class MainForm : Form
    {
        private string _currentLanguage;

        public MainForm()
        {
            InitializeComponent();
            InitializeMenu();
            LoadCurrentLanguage();
        }

        private void InitializeMenu()
        {
            // Create menu bar
            var menuStrip = new MenuStrip();
            
            // Add Language menu
            var languageMenu = new ToolStripMenuItem("Language");
            var selectLanguageItem = new ToolStripMenuItem("Select Language");
            selectLanguageItem.Click += SelectLanguageItem_Click;
            languageMenu.DropDownItems.Add(selectLanguageItem);
            
            menuStrip.Items.Add(languageMenu);
            Controls.Add(menuStrip);
            MainMenuStrip = menuStrip;
        }

        private void LoadCurrentLanguage()
        {
            try
            {
                _currentLanguage = AppSettings.Instance.LanguageCode;
                if (!LanguageManager.IsValidLanguageCode(_currentLanguage))
                {
                    _currentLanguage = "en"; // Fallback to English
                }
                UpdateUIForCurrentLanguage();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to load language settings", ex);
            }
        }

        private void SelectLanguageItem_Click(object sender, EventArgs e)
        {
            using (var languageForm = new LanguageSelectionForm())
            {
                if (languageForm.ShowDialog() == DialogResult.OK)
                {
                    ApplyLanguageChange();
                }
            }
        }

        private void ApplyLanguageChange()
        {
            try
            {
                var newLanguage = AppSettings.Instance.LanguageCode;
                if (newLanguage != _currentLanguage)
                {
                    _currentLanguage = newLanguage;
                    UpdateUIForCurrentLanguage();
                    MessageBox.Show(
                        $"Language changed to {CultureInfo.GetCultureInfo(_currentLanguage).DisplayName}",
                        "Language Changed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to apply language change", ex);
            }
        }

        private void UpdateUIForCurrentLanguage()
        {
            try
            {
                // Update form title
                Text = $"TextForge ({CultureInfo.GetCultureInfo(_currentLanguage).DisplayName})";
                
                // TODO: Update other UI elements as needed
                // This could include menus, labels, buttons, etc.
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to update UI for language", ex);
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Text = "TextForge";
            Size = new Size(800, 600);
            StartPosition = FormStartPosition.CenterScreen;
        }
    }
}
