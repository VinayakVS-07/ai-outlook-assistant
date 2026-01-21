namespace OutlookAddIn1
{
    partial class Ribbon2 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon2()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tabAIAssist = this.Factory.CreateRibbonTab();
            this.groupAITools = this.Factory.CreateRibbonGroup();
            this.splitCreateMail = this.Factory.CreateRibbonSplitButton();
            this.btnCreateProfessional = this.Factory.CreateRibbonButton();
            this.btnCreateCasual = this.Factory.CreateRibbonButton();
            this.btnSpellCheck = this.Factory.CreateRibbonButton();
            this.btnlanguage = this.Factory.CreateRibbonButton();
            this.btnReplyAssist = this.Factory.CreateRibbonButton();
            this.tabAIAssist.SuspendLayout();
            this.groupAITools.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabAIAssist
            // 
            this.tabAIAssist.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabAIAssist.Groups.Add(this.groupAITools);
            this.tabAIAssist.Label = "AI Assist";
            this.tabAIAssist.Name = "tabAIAssist";
            // 
            // groupAITools
            // 
            this.groupAITools.DialogLauncher = ribbonDialogLauncherImpl1;
            this.groupAITools.Items.Add(this.splitCreateMail);
            this.groupAITools.Items.Add(this.btnSpellCheck);
            this.groupAITools.Items.Add(this.btnlanguage);
            this.groupAITools.Items.Add(this.btnReplyAssist);
            this.groupAITools.Label = "AI Tools";
            this.groupAITools.Name = "groupAITools";
            // 
            // splitCreateMail
            // 
            this.splitCreateMail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitCreateMail.Image = global::OutlookAddIn1.Properties.Resources._7024891;
            this.splitCreateMail.Items.Add(this.btnCreateProfessional);
            this.splitCreateMail.Items.Add(this.btnCreateCasual);
            this.splitCreateMail.Label = "Create Mail";
            this.splitCreateMail.Name = "splitCreateMail";
            this.splitCreateMail.ScreenTip = "To create a mail";
            // 
            // btnCreateProfessional
            // 
            this.btnCreateProfessional.Image = global::OutlookAddIn1.Properties.Resources.Screenshot_2025_07_10_135559;
            this.btnCreateProfessional.Label = "Professional";
            this.btnCreateProfessional.Name = "btnCreateProfessional";
            this.btnCreateProfessional.ShowImage = true;
            this.btnCreateProfessional.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateProfessional_Click);
            // 
            // btnCreateCasual
            // 
            this.btnCreateCasual.Image = global::OutlookAddIn1.Properties.Resources.Screenshot_2025_07_10_135742;
            this.btnCreateCasual.Label = "Casual";
            this.btnCreateCasual.Name = "btnCreateCasual";
            this.btnCreateCasual.ShowImage = true;
            this.btnCreateCasual.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateCasual_Click);
            // 
            // btnSpellCheck
            // 
            this.btnSpellCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSpellCheck.Image = global::OutlookAddIn1.Properties.Resources.Screenshot_2025_07_08_220851;
            this.btnSpellCheck.Label = "Spell Check";
            this.btnSpellCheck.Name = "btnSpellCheck";
            this.btnSpellCheck.ScreenTip = "Check spelling";
            this.btnSpellCheck.ShowImage = true;
            this.btnSpellCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSpellCheck_Click);
            // 
            // btnlanguage
            // 
            this.btnlanguage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnlanguage.Image = global::OutlookAddIn1.Properties.Resources.stellar_mail_converter_logo;
            this.btnlanguage.Label = "Language Conversion";
            this.btnlanguage.Name = "btnlanguage";
            this.btnlanguage.ScreenTip = "To convert language";
            this.btnlanguage.ShowImage = true;
            this.btnlanguage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnlanguage_Click);
            // 
            // btnReplyAssist
            // 
            this.btnReplyAssist.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReplyAssist.Image = global::OutlookAddIn1.Properties.Resources._4627564;
            this.btnReplyAssist.Label = "Reply Assist";
            this.btnReplyAssist.Name = "btnReplyAssist";
            this.btnReplyAssist.ShowImage = true;
            this.btnReplyAssist.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplyAssist_Click);
            // 
            // Ribbon2
            // 
            this.Name = "Ribbon2";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tabAIAssist);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon2_Load);
            this.tabAIAssist.ResumeLayout(false);
            this.tabAIAssist.PerformLayout();
            this.groupAITools.ResumeLayout(false);
            this.groupAITools.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAIAssist;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAITools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSpellCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnlanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitCreateMail;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateProfessional;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateCasual;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplyAssist;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon2 Ribbon2
        {
            get { return this.GetRibbon<Ribbon2>(); }
        }
    }
}
