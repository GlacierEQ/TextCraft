﻿using System;
using System.ClientModel;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System.Configuration;
using PDFTron.WebViewer;
using PDFTron.WebViewer.Controls;
using PDFTron.WebViewer.Tools;
using TextForge.Services;

using TextForge.Services;

using TextForge.Services;

using TextForge.Services;


using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace TextForge
{
    public partial class Forge
    {
        private readonly ErrorHandlingService _errorHandler;

        public Forge(ErrorHandlingService errorHandler)
        {
            _errorHandler = errorHandler;
        }

    {
        private readonly ErrorHandlingService _errorHandler;

        public Forge(ErrorHandlingService errorHandler)
        {
            _errorHandler = errorHandler;
        }

    {
        private readonly ErrorHandlingService _errorHandler;

        public Forge(ErrorHandlingService errorHandler)
        {
            _errorHandler = errorHandler;
        }

    {
        private readonly ErrorHandlingService _errorHandler;

        public Forge(ErrorHandlingService errorHandler)
        {
            _errorHandler = errorHandler;
        }

    {
        // Public
        public static SystemChatMessage CommentSystemPrompt;
        public static readonly CultureLocalizationHelper CultureHelper = new CultureLocalizationHelper("TextForge.Forge", typeof(Forge).Assembly);
        public static readonly object InitializeDoor = new object();

        // Private
        private AboutBox _box;
        private static RibbonGroup _optionsBox;

        // Legal Document Processing
        private static readonly Dictionary<string, string> LegalPrompts = new Dictionary<string, string>
        {
            {"Divorce", "Analyze this divorce case document and identify key elements..."},
            {"Custody", "Review this child custody agreement and highlight important clauses..."},
            {"Labor", "Examine this labor dispute document and extract relevant legal arguments..."},
            {"Malpractice", "Analyze this legal malpractice case and identify potential issues..."},
            {"Defamation", "Review this defamation case document and extract key arguments..."}
        };

        private void Forge_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                    ThisAddIn.AddTaskPanes(Globals.ThisAddIn.Application.ActiveDocument);

                Thread startup = new Thread(InitializeForge);
                startup.SetApartmentState(ApartmentState.STA);
                startup.Start();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InitializeForge()
        {
            try
            {
                lock (InitializeDoor)
                {
                    if (!ThisAddIn.IsAddinInitialized)
                        ThisAddIn.InitializeAddIn();
                    
                    CommentSystemPrompt = new SystemChatMessage(ThisAddIn.SystemPromptLocalization["this.CommentSystemPrompt"]);

                    PopulateDropdownList(ThisAddIn.LanguageModelList);
                    InitializeLegalProcessing();
                }
                _box = new AboutBox();
                _optionsBox = this.OptionsGroup;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private void InitializeLegalProcessing()
        {
            var legalGroup = Globals.Factory.GetRibbonFactory().CreateRibbonGroup();
            legalGroup.Label = "Legal Tools";
            
            var divorceButton = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            divorceButton.Label = "Divorce Case";
            divorceButton.Click += async (s, e) => await ProcessLegalDocumentAsync("Divorce");
            legalGroup.Items.Add(divorceButton);

            var custodyButton = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            custodyButton.Label = "Custody Case";
            custodyButton.Click += async (s, e) => await ProcessLegalDocumentAsync("Custody");
            legalGroup.Items.Add(custodyButton);

            var laborButton = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            laborButton.Label = "Labor Dispute";
            laborButton.Click += async (s, e) => await ProcessLegalDocumentAsync("Labor");
            legalGroup.Items.Add(laborButton);

            var malpracticeButton = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            malpracticeButton.Label = "Malpractice";
            malpracticeButton.Click += async (s, e) => await ProcessLegalDocumentAsync("Malpractice");
            legalGroup.Items.Add(malpracticeButton);

            var defamationButton = Globals.Factory.GetRibbonFactory().CreateRibbonButton();
            defamationButton.Label = "Defamation";
            defamationButton.Click += async (s, e) => await ProcessLegalDocumentAsync("Defamation");
            legalGroup.Items.Add(defamationButton);

            this.Tabs[0].Groups.Add(legalGroup);
        }

        private async Task ProcessLegalDocumentAsync(string caseType, CancellationToken cancellationToken = default)
        {
            try
            {
                var document = Globals.ThisAddIn.Application.ActiveDocument;
                var range = document.Content;
                var prompt = LegalPrompts[caseType];

                await AnalyzeLegalTextAsync(prompt, range, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                // Operation was canceled, no need to display error
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private WebViewerControl _pdfViewer;
        
        private void InitializePDFViewer()
        {
            _pdfViewer = new WebViewerControl();
            _pdfViewer.Initialize(ConfigurationManager.AppSettings["PDFTronLicenseKey"]);
            _pdfViewer.Dock = DockStyle.Fill;
            
            var ocrButton = new ToolStripButton("OCR");
            ocrButton.Click += (s, e) => RunOCR();
            _pdfViewer.Toolbar.Items.Add(ocrButton);

            var aiAnalyzeButton = new ToolStripButton("AI Analyze");
            aiAnalyzeButton.Click += async (s, e) => await AnalyzeWithAIAsync();
            _pdfViewer.Toolbar.Items.Add(aiAnalyzeButton);
        }

        private void RunOCR()
        {
            try
            {
                var ocrOptions = new OCRModuleOptions
                {
                    Language = "eng",
                    OutputType = OCROutputType.SearchablePDF
                };
                _pdfViewer.Document.ApplyOCR(ocrOptions);
                MessageBox.Show("OCR completed successfully!");
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async Task AnalyzeWithAIAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                var text = _pdfViewer.Document.GetText();
                var analysis = await PerformAIAnalysisAsync(text, cancellationToken);
                ShowAnalysisResults(analysis);
            }
            catch (OperationCanceledException)
            {
                // Operation was canceled, no need to display error
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
            }
        }

        private async Task<string> PerformAIAnalysisAsync(string text, CancellationToken cancellationToken = default)
        {
            // Implement AI analysis logic here
            return await Task.FromResult("AI Analysis Results");
        }

        private void ShowAnalysisResults(string results)
        {
            var resultWindow = new Form
            {
                Text = "AI Analysis Results",
                Size = new System.Drawing.Size(800, 600)
            };
            
            var textBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Text = results
            };
            
            resultWindow.Controls.Add(textBox);
            resultWindow.Show();
        }

        private async Task AnalyzeLegalTextAsync(string systemPrompt, Word.Range range, CancellationToken cancellationToken = default)
        {
            try
            {
                var selectedText = range.Text;
                
                var pdfBytes = await ConvertWordToPDFAsync(range, cancellationToken);
                _pdfViewer.LoadDocument(pdfBytes);
                
                var annotationManager = _pdfViewer.GetAnnotationManager();
                annotationManager.AddTextAnnotation(selectedText, new System.Drawing.Point(10, 10));
                
                var modifiedWordContent = await ConvertPDFToWordAsync(pdfBytes, cancellationToken);
                
                range.Delete();
                range.Text = modifiedWordContent;
            }
            catch (OperationCanceledException)
            {
                // Operation was canceled, no need to display error
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
                range.Text = "Error processing document. Please try again.";
            }
        }

        private async Task<byte[]> ConvertWordToPDFAsync(Word.Range range, CancellationToken cancellationToken = default)
        {
            try
            {
                var tempFilePath = Path.GetTempFileName();
                var document = range.Document;
                document.SaveAs2(tempFilePath, Word.WdSaveFormat.wdFormatPDF);
                
                var pdfBytes = await File.ReadAllBytesAsync(tempFilePath, cancellationToken);
                File.Delete(tempFilePath);
                
                return pdfBytes;
            }
            catch (OperationCanceledException)
            {
                return Array.Empty<byte>();
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
                return Array.Empty<byte>();
            }
        }

        private async Task<string> ConvertPDFToWordAsync(byte[] pdfBytes, CancellationToken cancellationToken = default)
        {
            try
            {
                var tempFilePath = Path.GetTempFileName();
                await File.WriteAllBytesAsync(tempFilePath, pdfBytes, cancellationToken);
                
                var document = Globals.ThisAddIn.Application.Documents.Open(tempFilePath);
                var content = document.Content.Text;
                
                document.Close(false);
                File.Delete(tempFilePath);
                
                return content;
            }
            catch (OperationCanceledException)
            {
                return string.Empty;
            }
            catch (Exception ex)
            {
                CommonUtils.DisplayError(ex);
                return "Error converting PDF to Word content.";
            }
        }
    }
}
