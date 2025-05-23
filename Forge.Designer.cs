﻿using System.Globalization;

namespace TextForge
{
    partial class Forge : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Forge()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            Localize();
        }

        private void Localize()
        {
            CultureLocalizationHelper helper = new CultureLocalizationHelper("TextForge.Forge", typeof(Forge).Assembly);
            this.AboutButton.Label = helper.GetLocalizedString("this.AboutButton.Label");
            this.AboutButton.ScreenTip = helper.GetLocalizedString("this.AboutButton.ScreenTip");
            this.CancelButton.Label = helper.GetLocalizedString("this.CancelButton.Label");
            this.CancelButton.ScreenTip = helper.GetLocalizedString("this.CancelButton.ScreenTip");
            this.DefaultCheckBox.Label = helper.GetLocalizedString("this.DefaultCheckBox.Label");
            this.DefaultCheckBox.SuperTip = helper.GetLocalizedString("this.DefaultCheckBox.SuperTip");
            this.ForgeTab.Label = helper.GetLocalizedString("this.ForgeTab.Label");
            this.GenerateButton.Label = helper.GetLocalizedString("this.GenerateButton.Label");
            this.GenerateButton.SuperTip = helper.GetLocalizedString("this.GenerateButton.SuperTip");
            this.InfoGroup.Label = helper.GetLocalizedString("this.InfoGroup.Label");
            this.ModelListDropDown.Label = helper.GetLocalizedString("this.ModelListDropDown.Label");
            this.ModelListDropDown.SuperTip = helper.GetLocalizedString("this.ModelListDropDown.SuperTip");
            this.OptionsGroup.Label = helper.GetLocalizedString("this.OptionsGroup.Label");
            this.ProofreadButton.Label = helper.GetLocalizedString("this.ProofreadButton.Label");
            this.ProofreadButton.SuperTip = helper.GetLocalizedString("this.ProofreadButton.SuperTip");
            this.RAGControlButton.Label = helper.GetLocalizedString("this.RAGControlButton.Label");
            this.RAGControlButton.SuperTip = helper.GetLocalizedString("this.RAGControlButton.SuperTip");
            this.ReviewButton.Label = helper.GetLocalizedString("this.ReviewButton.Label");
            this.ReviewButton.SuperTip = helper.GetLocalizedString("this.ReviewButton.SuperTip");
            this.RewriteButton.Label = helper.GetLocalizedString("this.RewriteButton.Label");
            this.RewriteButton.SuperTip = helper.GetLocalizedString("this.RewriteButton.SuperTip");
            this.SettingsGroup.Label = helper.GetLocalizedString("this.SettingsGroup.Label");
            this.ToolsGroup.Label = helper.GetLocalizedString("this.ToolsGroup.Label");
            this.WritingToolsGallery.Label = helper.GetLocalizedString("this.WritingToolsGallery.Label");
            this.WritingToolsGallery.SuperTip = helper.GetLocalizedString("this.WritingToolsGallery.SuperTip");
        }


        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ForgeTab = this.Factory.CreateRibbonTab();
            this.ToolsGroup = this.Factory.CreateRibbonGroup();
            this.GenerateButton = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.WritingToolsGallery = this.Factory.CreateRibbonGallery();
            this.ReviewButton = this.Factory.CreateRibbonButton();
            this.ProofreadButton = this.Factory.CreateRibbonButton();
            this.RewriteButton = this.Factory.CreateRibbonButton();
            this.SettingsGroup = this.Factory.CreateRibbonGroup();
            this.RAGControlButton = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.ModelListDropDown = this.Factory.CreateRibbonDropDown();
            this.DefaultCheckBox = this.Factory.CreateRibbonCheckBox();
            this.InfoGroup = this.Factory.CreateRibbonGroup();
            this.AboutButton = this.Factory.CreateRibbonButton();
            this.OptionsGroup = this.Factory.CreateRibbonGroup();
            this.CancelButton = this.Factory.CreateRibbonButton();
            this.ForgeTab.SuspendLayout();
            this.ToolsGroup.SuspendLayout();
            this.SettingsGroup.SuspendLayout();
            this.InfoGroup.SuspendLayout();
            this.OptionsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // ForgeTab
            // 
            this.ForgeTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ForgeTab.Groups.Add(this.ToolsGroup);
            this.ForgeTab.Groups.Add(this.SettingsGroup);
            this.ForgeTab.Groups.Add(this.InfoGroup);
            this.ForgeTab.Groups.Add(this.OptionsGroup);
            this.ForgeTab.Label = "Craft";
            this.ForgeTab.Name = "ForgeTab";
            // 
            // ToolsGroup
            // 
            this.ToolsGroup.Items.Add(this.GenerateButton);
            this.ToolsGroup.Items.Add(this.separator3);
            this.ToolsGroup.Items.Add(this.WritingToolsGallery);
            this.ToolsGroup.Label = "Tools";
            this.ToolsGroup.Name = "ToolsGroup";
            // 
            // GenerateButton
            // 
            this.GenerateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GenerateButton.Image = global::TextForge.Properties.Resources.pen_high_contrast;
            this.GenerateButton.Label = "Generate";
            this.GenerateButton.Name = "GenerateButton";
            this.GenerateButton.ShowImage = true;
            this.GenerateButton.SuperTip = "Creates tailored responses using the user\'s prompt and the current document\'s con" +
    "tent.";
            this.GenerateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateButton_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // WritingToolsGallery
            // 
            this.WritingToolsGallery.Buttons.Add(this.ReviewButton);
            this.WritingToolsGallery.Buttons.Add(this.ProofreadButton);
            this.WritingToolsGallery.Buttons.Add(this.RewriteButton);
            this.WritingToolsGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.WritingToolsGallery.Image = global::TextForge.Properties.Resources.memo_high_contrast;
            this.WritingToolsGallery.Label = "Writing Tools";
            this.WritingToolsGallery.Name = "WritingToolsGallery";
            this.WritingToolsGallery.ShowImage = true;
            this.WritingToolsGallery.SuperTip = "Enhance your writing with AI-powered tools for grammar, style, and clarity.";
            this.WritingToolsGallery.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WritingToolsGallery_ButtonClick);
            // 
            // ReviewButton
            // 
            this.ReviewButton.Image = global::TextForge.Properties.Resources.clipboard_high_contrast;
            this.ReviewButton.Label = "Review";
            this.ReviewButton.Name = "ReviewButton";
            this.ReviewButton.ShowImage = true;
            this.ReviewButton.SuperTip = "Enhance your document with AI-driven comments and suggestions.";
            // 
            // ProofreadButton
            // 
            this.ProofreadButton.Image = global::TextForge.Properties.Resources.face_with_monocle_high_contrast;
            this.ProofreadButton.Label = "Proofread";
            this.ProofreadButton.Name = "ProofreadButton";
            this.ProofreadButton.ShowImage = true;
            this.ProofreadButton.SuperTip = "Check for spelling, grammar, and style errors to polish your document.";
            // 
            // RewriteButton
            // 
            this.RewriteButton.Image = global::TextForge.Properties.Resources.counterclockwise_arrows_button_high_contrast;
            this.RewriteButton.Label = "Rewrite";
            this.RewriteButton.Name = "RewriteButton";
            this.RewriteButton.ShowImage = true;
            this.RewriteButton.SuperTip = "Revise and enhance your text with AI-powered suggestions.";
            // 
            // SettingsGroup
            // 
            this.SettingsGroup.Items.Add(this.RAGControlButton);
            this.SettingsGroup.Items.Add(this.separator2);
            this.SettingsGroup.Items.Add(this.ModelListDropDown);
            this.SettingsGroup.Items.Add(this.DefaultCheckBox);
            this.SettingsGroup.Label = "Settings";
            this.SettingsGroup.Name = "SettingsGroup";
            // 
            // RAGControlButton
            // 
            this.RAGControlButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RAGControlButton.Image = global::TextForge.Properties.Resources.gear_high_contrast;
            this.RAGControlButton.Label = "RAG Control";
            this.RAGControlButton.Name = "RAGControlButton";
            this.RAGControlButton.ShowImage = true;
            this.RAGControlButton.SuperTip = "Manage files within the Retrieval Augmentation Generation (RAG) system.";
            this.RAGControlButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RAGControlButton_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // ModelListDropDown
            // 
            this.ModelListDropDown.Label = "Model List";
            this.ModelListDropDown.Name = "ModelListDropDown";
            this.ModelListDropDown.ShowLabel = false;
            this.ModelListDropDown.SizeString = "XXXXXXXXXXXXXXXXXXXXXXXXX";
            this.ModelListDropDown.SuperTip = "Select a language model for use in TextCraft.";
            this.ModelListDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ModelListDropDown_SelectionChanged);
            // 
            // DefaultCheckBox
            // 
            this.DefaultCheckBox.Label = "Default";
            this.DefaultCheckBox.Name = "DefaultCheckBox";
            this.DefaultCheckBox.SuperTip = "Set the default language model for TextCraft.";
            this.DefaultCheckBox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DefaultCheckBox_Click);
            // 
            // InfoGroup
            // 
            this.InfoGroup.Items.Add(this.AboutButton);
            this.InfoGroup.Label = "Info";
            this.InfoGroup.Name = "InfoGroup";
            // 
            // AboutButton
            // 
            this.AboutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AboutButton.Image = global::TextForge.Properties.Resources.information_high_contrast;
            this.AboutButton.Label = "About";
            this.AboutButton.Name = "AboutButton";
            this.AboutButton.ScreenTip = "Provides details about this application.";
            this.AboutButton.ShowImage = true;
            this.AboutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AboutButton_Click);
            // 
            // OptionsGroup
            // 
            this.OptionsGroup.Items.Add(this.CancelButton);
            this.OptionsGroup.Label = "Options";
            this.OptionsGroup.Name = "OptionsGroup";
            this.OptionsGroup.Visible = false;
            // 
            // CancelButton
            // 
            this.CancelButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CancelButton.Image = global::TextForge.Properties.Resources.stop_sign_flat;
            this.CancelButton.Label = "Cancel";
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.ScreenTip = "Halts the text generation process.";
            this.CancelButton.ShowImage = true;
            this.CancelButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CancelButton_Click);
            // 
            // Forge
            // 
            this.Name = "Forge";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ForgeTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Forge_Load);
            this.ForgeTab.ResumeLayout(false);
            this.ForgeTab.PerformLayout();
            this.ToolsGroup.ResumeLayout(false);
            this.ToolsGroup.PerformLayout();
            this.SettingsGroup.ResumeLayout(false);
            this.SettingsGroup.PerformLayout();
            this.InfoGroup.ResumeLayout(false);
            this.InfoGroup.PerformLayout();
            this.OptionsGroup.ResumeLayout(false);
            this.OptionsGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ForgeTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SettingsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ModelListDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox DefaultCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RAGControlButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup InfoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AboutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CancelButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup OptionsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery WritingToolsGallery;
        private Microsoft.Office.Tools.Ribbon.RibbonButton ReviewButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton ProofreadButton;
        private Microsoft.Office.Tools.Ribbon.RibbonButton RewriteButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
    }

    partial class ThisRibbonCollection
    {
        internal Forge Forge
        {
            get { return this.GetRibbon<Forge>(); }
        }
    }
}
