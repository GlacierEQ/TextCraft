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
            var languageMenu = new ToolStripMenuItem(LocalizationHelper.GetString("Menu_Language"));
            var selectLanguageItem = new ToolStripMenuItem(LocalizationHelper.GetString("Menu_SelectLanguage"));
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
                        string.Format(LocalizationHelper.GetString("Message_LanguageChanged"), 
                            CultureInfo.GetCultureInfo(_currentLanguage).DisplayName),
                        LocalizationHelper.GetString("Menu_Language"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to apply language change", ex);
            }
        }

        private void UpdateFormTitle()
        {
            try
            {
                Text = $"{LocalizationHelper.GetString("FormTitle")} ({CultureInfo.GetCultureInfo(_currentLanguage).DisplayName})";
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to update form title", ex);
                Text = "TextForge";
            }
        }

        private void UpdateMenuItems()
        {
            try
            {
                var menuStrip = MainMenuStrip;
                if (menuStrip != null && menuStrip.Items.Count > 0)
                {
                    // Update main menu items
                    menuStrip.Items[0].Text = LocalizationHelper.GetString("Menu_File");
                    menuStrip.Items[1].Text = LocalizationHelper.GetString("Menu_Edit");
                    menuStrip.Items[2].Text = LocalizationHelper.GetString("Menu_View");
                    menuStrip.Items[3].Text = LocalizationHelper.GetString("Menu_Help");
                    
                    // Update language menu
                    var languageMenu = menuStrip.Items[3] as ToolStripMenuItem;
                    if (languageMenu != null && languageMenu.DropDownItems.Count > 0)
                    {
                        languageMenu.Text = LocalizationHelper.GetString("Menu_Language");
                        languageMenu.DropDownItems[0].Text = LocalizationHelper.GetString("Menu_SelectLanguage");
                    }
                }
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to update menu items", ex);
            }
        }

        private void UpdateUIForCurrentLanguage()
        {
            try
            {
                UpdateFormTitle();
                UpdateMenuItems();
                
                // Clear localization cache when language changes
                LocalizationHelper.ClearCache();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError("Failed to update UI for language", ex);
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            UpdateFormTitle();
            Size = new Size(800, 600);
            StartPosition = FormStartPosition.CenterScreen;
        }
    }
}
