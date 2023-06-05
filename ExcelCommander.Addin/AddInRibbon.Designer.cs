namespace ExcelCommander.Addin
{
    partial class AddInRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddInRibbon()
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
            this.startButton = this.Factory.CreateRibbonButton();
            this.stopButton = this.Factory.CreateRibbonButton();
            this.statusLabel = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
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
            this.group1.Items.Add(this.startButton);
            this.group1.Items.Add(this.stopButton);
            this.group1.Items.Add(this.statusLabel);
            this.group1.Label = "Excel Commander";
            this.group1.Name = "group1";
            // 
            // startButton
            // 
            this.startButton.Description = "Start the service server.";
            this.startButton.KeyTip = "SA";
            this.startButton.Label = "Start";
            this.startButton.Name = "startButton";
            this.startButton.SuperTip = "Start the service server.";
            this.startButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startButton_Click);
            // 
            // stopButton
            // 
            this.stopButton.Description = "Stop the service server.";
            this.stopButton.Enabled = false;
            this.stopButton.KeyTip = "ST";
            this.stopButton.Label = "Stop";
            this.stopButton.Name = "stopButton";
            this.stopButton.SuperTip = "Stop the service server.";
            this.stopButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.stopButton_Click);
            // 
            // statusLabel
            // 
            this.statusLabel.Label = "Ready.";
            this.statusLabel.Name = "statusLabel";
            // 
            // AddInRibbon
            // 
            this.Name = "AddInRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AddInRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton startButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton stopButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel statusLabel;
    }

    partial class ThisRibbonCollection
    {
        internal AddInRibbon AddInRibbon
        {
            get { return this.GetRibbon<AddInRibbon>(); }
        }
    }
}
