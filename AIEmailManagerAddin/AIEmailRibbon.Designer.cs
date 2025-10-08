namespace AIEmailManagerAddin
{
    partial class AIEmailRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AIEmailRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnAnalyzeCurrent = this.Factory.CreateRibbonButton();
            this.btnAnalyzeFolder = this.Factory.CreateRibbonButton();
            this.btnRefreshEmails = this.Factory.CreateRibbonButton();
            this.groupMeetings = this.Factory.CreateRibbonGroup();
            this.btnAnalyzeMeeting = this.Factory.CreateRibbonButton();
            this.btnAnalyzeMeetings = this.Factory.CreateRibbonButton();
            this.btnRefreshMeetings = this.Factory.CreateRibbonButton();
            this.groupSystem = this.Factory.CreateRibbonGroup();
            this.btnStats = this.Factory.CreateRibbonButton();
            this.btnOpenWeb = this.Factory.CreateRibbonButton();
            this.btnLearningManagement = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.groupMeetings.SuspendLayout();
            this.groupSystem.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.groupMeetings);
            this.tab2.Groups.Add(this.groupSystem);
            this.tab2.Label = "ניהול עם AI";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnAnalyzeCurrent);
            this.group2.Items.Add(this.btnAnalyzeFolder);
            this.group2.Items.Add(this.btnRefreshEmails);
            this.group2.Label = "ניהול מיילים";
            this.group2.Name = "group2";
            // 
            // btnAnalyzeCurrent
            // 
            this.btnAnalyzeCurrent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnalyzeCurrent.Label = "נתח מייל נוכחי";
            this.btnAnalyzeCurrent.Name = "btnAnalyzeCurrent";
            this.btnAnalyzeCurrent.OfficeImageId = "FindDialog";
            this.btnAnalyzeCurrent.ShowImage = true;
            this.btnAnalyzeCurrent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnalyzeCurrent_Click);
            // 
            // btnAnalyzeFolder
            // 
            this.btnAnalyzeFolder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnalyzeFolder.Label = "נתח מיילים נבחרים";
            this.btnAnalyzeFolder.Name = "btnAnalyzeFolder";
            this.btnAnalyzeFolder.OfficeImageId = "MailMergeRecipientsEditList";
            this.btnAnalyzeFolder.ShowImage = true;
            this.btnAnalyzeFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnalyzeFolder_Click);
            // 
            // btnRefreshEmails
            // 
            this.btnRefreshEmails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefreshEmails.Label = "רענן מיילים";
            this.btnRefreshEmails.Name = "btnRefreshEmails";
            this.btnRefreshEmails.OfficeImageId = "Refresh";
            this.btnRefreshEmails.ShowImage = true;
            this.btnRefreshEmails.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefreshEmails_Click);
            // 
            // groupMeetings
            // 
            this.groupMeetings.Items.Add(this.btnAnalyzeMeeting);
            this.groupMeetings.Items.Add(this.btnAnalyzeMeetings);
            this.groupMeetings.Items.Add(this.btnRefreshMeetings);
            this.groupMeetings.Label = "ניהול פגישות";
            this.groupMeetings.Name = "groupMeetings";
            // 
            // btnAnalyzeMeeting
            // 
            this.btnAnalyzeMeeting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnalyzeMeeting.Label = "נתח פגישה נוכחית";
            this.btnAnalyzeMeeting.Name = "btnAnalyzeMeeting";
            this.btnAnalyzeMeeting.OfficeImageId = "CalendarInsert";
            this.btnAnalyzeMeeting.ShowImage = true;
            this.btnAnalyzeMeeting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnalyzeMeeting_Click);
            // 
            // btnAnalyzeMeetings
            // 
            this.btnAnalyzeMeetings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnalyzeMeetings.Label = "נתח פגישות נבחרות";
            this.btnAnalyzeMeetings.Name = "btnAnalyzeMeetings";
            this.btnAnalyzeMeetings.OfficeImageId = "RecurrenceEdit";
            this.btnAnalyzeMeetings.ShowImage = true;
            this.btnAnalyzeMeetings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnalyzeMeetings_Click);
            // 
            // btnRefreshMeetings
            // 
            this.btnRefreshMeetings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefreshMeetings.Label = "רענן פגישות";
            this.btnRefreshMeetings.Name = "btnRefreshMeetings";
            this.btnRefreshMeetings.OfficeImageId = "Refresh";
            this.btnRefreshMeetings.ShowImage = true;
            this.btnRefreshMeetings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefreshMeetings_Click);
            // 
            // groupSystem
            // 
            this.groupSystem.Items.Add(this.btnStats);
            this.groupSystem.Items.Add(this.btnOpenWeb);
            this.groupSystem.Items.Add(this.btnLearningManagement);
            this.groupSystem.Items.Add(this.btnSettings);
            this.groupSystem.Label = "ניהול מערכת";
            this.groupSystem.Name = "groupSystem";
            // 
            // btnStats
            // 
            this.btnStats.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStats.Label = "ניהול מיילים";
            this.btnStats.Name = "btnStats";
            this.btnStats.OfficeImageId = "ReadingPaneRight";
            this.btnStats.ShowImage = true;
            this.btnStats.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStats_Click);
            // 
            // btnOpenWeb
            // 
            this.btnOpenWeb.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenWeb.Label = "ניהול פגישות";
            this.btnOpenWeb.Name = "btnOpenWeb";
            this.btnOpenWeb.OfficeImageId = "DateAndTimeInsert";
            this.btnOpenWeb.ShowImage = true;
            this.btnOpenWeb.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenWeb_Click);
            // 
            // btnLearningManagement
            // 
            this.btnLearningManagement.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLearningManagement.Label = "ניהול למידה";
            this.btnLearningManagement.Name = "btnLearningManagement";
            this.btnLearningManagement.OfficeImageId = "ReviewAcceptChange";
            this.btnLearningManagement.ShowImage = true;
            this.btnLearningManagement.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLearningManagement_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Label = "קונסול";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.OfficeImageId = "ViewCode";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            // 
            // AIEmailRibbon
            // 
            this.Name = "AIEmailRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.groupMeetings.ResumeLayout(false);
            this.groupMeetings.PerformLayout();
            this.groupSystem.ResumeLayout(false);
            this.groupSystem.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAnalyzeCurrent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAnalyzeFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefreshEmails;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMeetings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAnalyzeMeeting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAnalyzeMeetings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefreshMeetings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSystem;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLearningManagement;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStats;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenWeb;
    }

    partial class ThisRibbonCollection
    {
        internal AIEmailRibbon Ribbon1
        {
            get { return this.GetRibbon<AIEmailRibbon>(); }
        }
    }
}
