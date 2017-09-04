namespace SynchronizitätsUntersuchung
{
    partial class SynchronizitaetUntersuchung : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SynchronizitaetUntersuchung()
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
            this.synchronizitaetsUntersuchungGroup = this.Factory.CreateRibbonGroup();
            this.btnUntersuchungStarten = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.synchronizitaetsUntersuchungGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.synchronizitaetsUntersuchungGroup);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // synchronizitaetsUntersuchungGroup
            // 
            this.synchronizitaetsUntersuchungGroup.Items.Add(this.btnUntersuchungStarten);
            this.synchronizitaetsUntersuchungGroup.Label = "Synchronizität untersuchen";
            this.synchronizitaetsUntersuchungGroup.Name = "synchronizitaetsUntersuchungGroup";
            this.synchronizitaetsUntersuchungGroup.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupContactFind");
            // 
            // btnUntersuchungStarten
            // 
            this.btnUntersuchungStarten.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUntersuchungStarten.Label = "Untersuchung durchführen";
            this.btnUntersuchungStarten.Name = "btnUntersuchungStarten";
            this.btnUntersuchungStarten.ShowImage = true;
            this.btnUntersuchungStarten.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUntersuchungStarten_Click);
            // 
            // SynchronizitaetUntersuchung
            // 
            this.Name = "SynchronizitaetUntersuchung";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SynchronizitaetUntersuchungRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.synchronizitaetsUntersuchungGroup.ResumeLayout(false);
            this.synchronizitaetsUntersuchungGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup synchronizitaetsUntersuchungGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUntersuchungStarten;
    }

    partial class ThisRibbonCollection
    {
        internal SynchronizitaetUntersuchung SynchronizitaetUntersuchungRibbon
        {
            get { return this.GetRibbon<SynchronizitaetUntersuchung>(); }
        }
    }
}
