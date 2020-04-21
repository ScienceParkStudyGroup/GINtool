namespace GINtool
{
    partial class GinRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public GinRibbon()
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
            this.TabGINtool = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btApply = this.Factory.CreateRibbonButton();
            this.tglDense = this.Factory.CreateRibbonToggleButton();
            this.btLoad = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btSelectFile = this.Factory.CreateRibbonButton();
            this.lbRefFileName = this.Factory.CreateRibbonLabel();
            this.btRegDirMap = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.ddRegulon = this.Factory.CreateRibbonDropDown();
            this.ddBSU = this.Factory.CreateRibbonDropDown();
            this.ddDir = this.Factory.CreateRibbonDropDown();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ebLow = this.Factory.CreateRibbonEditBox();
            this.ebMid = this.Factory.CreateRibbonEditBox();
            this.ebHigh = this.Factory.CreateRibbonEditBox();
            this.TabGINtool.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabGINtool
            // 
            this.TabGINtool.Groups.Add(this.group1);
            this.TabGINtool.Groups.Add(this.group2);
            this.TabGINtool.Label = "GIN tool";
            this.TabGINtool.Name = "TabGINtool";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btApply);
            this.group1.Items.Add(this.tglDense);
            this.group1.Items.Add(this.btLoad);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btSelectFile);
            this.group1.Items.Add(this.lbRefFileName);
            this.group1.Items.Add(this.btRegDirMap);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.ddRegulon);
            this.group1.Items.Add(this.ddBSU);
            this.group1.Items.Add(this.ddDir);
            this.group1.Label = "GIN tool";
            this.group1.Name = "group1";
            // 
            // btApply
            // 
            this.btApply.Enabled = false;
            this.btApply.Image = global::GINtool.Properties.Resources.ApplyCodeChanges_16x;
            this.btApply.Label = "apply to range";
            this.btApply.Name = "btApply";
            this.btApply.ScreenTip = "apply the mapping to the selected range of cells";
            this.btApply.ShowImage = true;
            this.btApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btApply_Click);
            // 
            // tglDense
            // 
            this.tglDense.Enabled = false;
            this.tglDense.Image = global::GINtool.Properties.Resources.Span_16x;
            this.tglDense.Label = "dense/sparse output";
            this.tglDense.Name = "tglDense";
            this.tglDense.ScreenTip = "select the condensed (default) or sparse output mode";
            this.tglDense.ShowImage = true;
            this.tglDense.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglDense_Click);
            // 
            // btLoad
            // 
            this.btLoad.Enabled = false;
            this.btLoad.Image = global::GINtool.Properties.Resources.Refetch_16x;
            this.btLoad.Label = "load reference data";
            this.btLoad.Name = "btLoad";
            this.btLoad.ScreenTip = "load reference data into memory";
            this.btLoad.ShowImage = true;
            this.btLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btLoad_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btSelectFile
            // 
            this.btSelectFile.Image = global::GINtool.Properties.Resources.FileSystemDriverFile_16x;
            this.btSelectFile.Label = "set reference file";
            this.btSelectFile.Name = "btSelectFile";
            this.btSelectFile.ScreenTip = "select the reference file that contains the mapping tables";
            this.btSelectFile.ShowImage = true;
            this.btSelectFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btSelectFile_Click);
            // 
            // lbRefFileName
            // 
            this.lbRefFileName.Label = "reference file location";
            this.lbRefFileName.Name = "lbRefFileName";
            this.lbRefFileName.ScreenTip = "placeholder for the selected reference file";
            // 
            // btRegDirMap
            // 
            this.btRegDirMap.Enabled = false;
            this.btRegDirMap.Image = global::GINtool.Properties.Resources.StatisticsUp_16x;
            this.btRegDirMap.Label = "regulon direction map";
            this.btRegDirMap.Name = "btRegDirMap";
            this.btRegDirMap.ScreenTip = "define the text mappings for the direction of the regulons";
            this.btRegDirMap.ShowImage = true;
            this.btRegDirMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btRegDirMap_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // ddRegulon
            // 
            this.ddRegulon.Enabled = false;
            this.ddRegulon.Image = global::GINtool.Properties.Resources.Driver_16x;
            this.ddRegulon.Label = "regulon column";
            this.ddRegulon.Name = "ddRegulon";
            this.ddRegulon.ScreenTip = "specify the column that contains the regulator names";
            this.ddRegulon.ShowImage = true;
            this.ddRegulon.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddRegulon_SelectionChanged);
            // 
            // ddBSU
            // 
            this.ddBSU.Enabled = false;
            this.ddBSU.Image = global::GINtool.Properties.Resources.Target_16x;
            this.ddBSU.Label = "bsu column";
            this.ddBSU.Name = "ddBSU";
            this.ddBSU.ScreenTip = "specify the column that contains the bsu code";
            this.ddBSU.ShowImage = true;
            this.ddBSU.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddBSU_SelectionChanged);
            // 
            // ddDir
            // 
            this.ddDir.Image = global::GINtool.Properties.Resources.DownloadLog_16x;
            this.ddDir.Label = "direction column";
            this.ddDir.Name = "ddDir";
            this.ddDir.ShowImage = true;
            this.ddDir.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddDir_SelectionChanged);
            // 
            // group2
            // 
            this.group2.Items.Add(this.ebLow);
            this.group2.Items.Add(this.ebMid);
            this.group2.Items.Add(this.ebHigh);
            this.group2.Label = "fc ranges";
            this.group2.Name = "group2";
            // 
            // ebLow
            // 
            this.ebLow.Label = "low";
            this.ebLow.Name = "ebLow";
            this.ebLow.Text = null;
            this.ebLow.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ebLow_TextChanged);
            // 
            // ebMid
            // 
            this.ebMid.Label = "mid";
            this.ebMid.Name = "ebMid";
            this.ebMid.Text = null;
            // 
            // ebHigh
            // 
            this.ebHigh.Label = "high";
            this.ebHigh.Name = "ebHigh";
            this.ebHigh.Text = null;
            // 
            // GinRibbon
            // 
            this.Name = "GinRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.TabGINtool);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.GinRibbon_Load);
            this.TabGINtool.ResumeLayout(false);
            this.TabGINtool.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabGINtool;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btApply;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tglDense;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btSelectFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbRefFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddRegulon;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddBSU;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btRegDirMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddDir;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebLow;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebMid;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebHigh;
    }

    partial class ThisRibbonCollection
    {
        internal GinRibbon GinRibbon
        {
            get { return this.GetRibbon<GinRibbon>(); }
        }
    }
}
