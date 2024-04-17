namespace VisioAddinPersistingShapes
{
    partial class AddinRibbonComponent : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddinRibbonComponent()
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
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.Group1 = this.Factory.CreateRibbonGroup();
            this.btnListFormats = this.Factory.CreateRibbonButton();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.btnLoad = this.Factory.CreateRibbonButton();
            this.ddFormat = this.Factory.CreateRibbonDropDown();
            this.Tab1.SuspendLayout();
            this.Group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1
            // 
            this.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab1.ControlId.OfficeId = "TabHome";
            this.Tab1.Groups.Add(this.Group1);
            this.Tab1.Label = "TabHome";
            this.Tab1.Name = "Tab1";
            // 
            // Group1
            // 
            this.Group1.Items.Add(this.btnListFormats);
            this.Group1.Items.Add(this.btnSave);
            this.Group1.Items.Add(this.btnLoad);
            this.Group1.Items.Add(this.ddFormat);
            this.Group1.Label = "VisioAddinPersistingShapes";
            this.Group1.Name = "Group1";
            // 
            // btnListFormats
            // 
            this.btnListFormats.Label = "Get Formats";
            this.btnListFormats.Name = "btnListFormats";
            this.btnListFormats.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnListFormats_Click);
            // 
            // btnSave
            // 
            this.btnSave.Label = "Save to Clipboard";
            this.btnSave.Name = "btnSave";
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Label = "Load from Clipboard";
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoad_Click);
            // 
            // ddFormat
            // 
            this.ddFormat.Label = "Format";
            this.ddFormat.Name = "ddFormat";
            // 
            // AddinRibbonComponent
            // 
            this.Name = "AddinRibbonComponent";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.Tab1);
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.Group1.ResumeLayout(false);
            this.Group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnListFormats;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddFormat;
    }

    partial class ThisRibbonCollection
    {
        internal AddinRibbonComponent Ribbon
        {
            get { return this.GetRibbon<AddinRibbonComponent>(); }
        }
    }
}
