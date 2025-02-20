using System;
using System.Windows.Forms;

namespace TextForge
{
    public partial class MainForm : Form
    {
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
            // TODO: Load saved language from settings
            // Update UI with current language
        }

        private void SelectLanguageItem_Click(object sender, EventArgs e)
        {
            using (var languageForm = new LanguageSelectionForm())
            {
                if (languageForm.ShowDialog() == DialogResult.OK)
                {
                    // Handle language change
                    ApplyLanguageChange();
                }
            }
        }

        private void ApplyLanguageChange()
        {
            // TODO: Implement language change logic
            // 1. Update UI elements
            // 2. Save setting
            // 3. Refresh any language-dependent content
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
