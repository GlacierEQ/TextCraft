using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace TextForge
{
    public partial class LanguageSelectionForm : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly ComboBox _languageComboBox;
        private readonly Button _downloadButton;
        private readonly Button _applyButton;

        public LanguageSelectionForm()
        {
            InitializeComponent();
            
            // Form setup
            Text = LocalizationHelper.GetString("Form_LanguageSelection");
            Size = new Size(400, 200);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            // Language ComboBox
            _languageComboBox = new ComboBox
            {
                Location = new Point(20, 20),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Download Button
            _downloadButton = new Button
            {
                Text = LocalizationHelper.GetString("Button_Download"),
                Location = new Point(230, 19),
                Width = 100
            };
            _downloadButton.Click += DownloadButton_Click;

            // Progress Bar
            _progressBar = new ProgressBar
            {
                Location = new Point(20, 60),
                Width = 310,
                Visible = false
            };

            // Apply Button
            _applyButton = new Button
            {
                Text = LocalizationHelper.GetString("Button_Apply"),
                Location = new Point(20, 100),
                Width = 100
            };
            _applyButton.Click += ApplyButton_Click;

            // Add controls to form
            Controls.Add(_languageComboBox);
            Controls.Add(_downloadButton);
            Controls.Add(_progressBar);
            Controls.Add(_applyButton);

            LoadLanguages();
        }

        private void LoadLanguages()
        {
            _languageComboBox.Items.Clear();
            var languages = LanguageManager.GetAvailableLanguages();
            
            foreach (var lang in languages)
            {
                _languageComboBox.Items.Add(new LanguageItem(lang));
            }

            _languageComboBox.SelectedIndex = 0;
        }

        private async void DownloadButton_Click(object sender, EventArgs e)
        {
            try
            {
                _downloadButton.Enabled = false;
                _progressBar.Visible = true;

                var selectedLang = _languageComboBox.SelectedItem as LanguageItem;
                if (selectedLang != null)
                {
                    var progress = new Progress<int>(percent =>
                    {
                        _progressBar.Value = percent;
                    });

                    await LanguageManager.DownloadLanguageAsync(selectedLang.Code, progress);
                    LoadLanguages();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to download language: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _downloadButton.Enabled = true;
                _progressBar.Visible = false;
            }
        }

        private void ApplyButton_Click(object sender, EventArgs e)
        {
            var selectedLang = _languageComboBox.SelectedItem as LanguageItem;
            if (selectedLang != null)
            {
                // TODO: Implement language change in main application
                MessageBox.Show($"Language changed to {selectedLang.DisplayName}", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        private class LanguageItem
        {
            public string Code { get; }
            public string DisplayName { get; }

            public LanguageItem(string code)
            {
                Code = code;
                DisplayName = CultureInfo.GetCultureInfo(code).DisplayName;
            }

            public override string ToString() => DisplayName;
        }
    }
}
