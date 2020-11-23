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
            this.splitButton3 = this.Factory.CreateRibbonSplitButton();
            this.but_pvalues = this.Factory.CreateRibbonButton();
            this.but_fc = this.Factory.CreateRibbonButton();
            this.btLoad = this.Factory.CreateRibbonButton();
            this.splitbtnEA = this.Factory.CreateRibbonSplitButton();
            this.cbComposit = this.Factory.CreateRibbonCheckBox();
            this.cbQplot = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.tglTaskPane = this.Factory.CreateRibbonToggleButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectRegulonFile = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnRegulonFileName = this.Factory.CreateRibbonButton();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectOperonFile = this.Factory.CreateRibbonButton();
            this.btnResetOperonFile = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.btnOperonFile = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.ddBSU = this.Factory.CreateRibbonDropDown();
            this.ddRegulon = this.Factory.CreateRibbonDropDown();
            this.ddGene = this.Factory.CreateRibbonDropDown();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.ddDir = this.Factory.CreateRibbonDropDown();
            this.btRegDirMap = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ebLow = this.Factory.CreateRibbonEditBox();
            this.ebMid = this.Factory.CreateRibbonEditBox();
            this.ebHigh = this.Factory.CreateRibbonEditBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.editMinPval = this.Factory.CreateRibbonEditBox();
            this.grpPref = this.Factory.CreateRibbonGroup();
            this.btnPalette = this.Factory.CreateRibbonGallery();
            this.clrBerry = this.Factory.CreateRibbonButton();
            this.clrBright = this.Factory.CreateRibbonButton();
            this.clrBrightPastel = this.Factory.CreateRibbonButton();
            this.clrChocolate = this.Factory.CreateRibbonButton();
            this.clrExcel = this.Factory.CreateRibbonButton();
            this.clrGray = this.Factory.CreateRibbonButton();
            this.edtMaxGroups = this.Factory.CreateRibbonEditBox();
            this.TabGINtool.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group6.SuspendLayout();
            this.group5.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.grpPref.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabGINtool
            // 
            this.TabGINtool.Groups.Add(this.group1);
            this.TabGINtool.Groups.Add(this.group3);
            this.TabGINtool.Groups.Add(this.group6);
            this.TabGINtool.Groups.Add(this.group5);
            this.TabGINtool.Groups.Add(this.group2);
            this.TabGINtool.Groups.Add(this.group4);
            this.TabGINtool.Groups.Add(this.grpPref);
            this.TabGINtool.Label = "GIN tool";
            this.TabGINtool.Name = "TabGINtool";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btApply);
            this.group1.Items.Add(this.splitButton3);
            this.group1.Items.Add(this.btLoad);
            this.group1.Items.Add(this.splitbtnEA);
            this.group1.Items.Add(this.tglTaskPane);
            this.group1.Label = "main";
            this.group1.Name = "group1";
            // 
            // btApply
            // 
            this.btApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btApply.Image = global::GINtool.Properties.Resources.gear;
            this.btApply.Label = "apply to range";
            this.btApply.Name = "btApply";
            this.btApply.ScreenTip = "Start the analysis (PVALUE,FC,BSU)";
            this.btApply.ShowImage = true;
            this.btApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btApply_Click);
            // 
            // splitButton3
            // 
            this.splitButton3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton3.Image = global::GINtool.Properties.Resources.pvalue;
            this.splitButton3.Items.Add(this.but_pvalues);
            this.splitButton3.Items.Add(this.but_fc);
            this.splitButton3.Label = "use p values";
            this.splitButton3.Name = "splitButton3";
            this.splitButton3.ScreenTip = "Select which cut-off to use for the combined report";
            // 
            // but_pvalues
            // 
            this.but_pvalues.Image = global::GINtool.Properties.Resources.pvalue;
            this.but_pvalues.Label = "use p values";
            this.but_pvalues.Name = "but_pvalues";
            this.but_pvalues.ScreenTip = "Use the p-value cut-off for the combined report";
            this.but_pvalues.ShowImage = true;
            this.but_pvalues.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // but_fc
            // 
            this.but_fc.Image = global::GINtool.Properties.Resources.fold_change;
            this.but_fc.Label = "use fold changes";
            this.but_fc.Name = "but_fc";
            this.but_fc.ScreenTip = "Use the low FC value as cut-off for the combined report";
            this.but_fc.ShowImage = true;
            this.but_fc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.but_fc_Click);
            // 
            // btLoad
            // 
            this.btLoad.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btLoad.Enabled = false;
            this.btLoad.Image = global::GINtool.Properties.Resources.stack;
            this.btLoad.Label = "load reference data";
            this.btLoad.Name = "btLoad";
            this.btLoad.ScreenTip = "load reference data into memory";
            this.btLoad.ShowImage = true;
            this.btLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btLoad_Click);
            // 
            // splitbtnEA
            // 
            this.splitbtnEA.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitbtnEA.Image = global::GINtool.Properties.Resources.barchart;
            this.splitbtnEA.Items.Add(this.cbComposit);
            this.splitbtnEA.Items.Add(this.cbQplot);
            this.splitbtnEA.Items.Add(this.separator1);
            this.splitbtnEA.Label = "enrichment plots";
            this.splitbtnEA.Name = "splitbtnEA";
            // 
            // cbComposit
            // 
            this.cbComposit.Label = "composit plot";
            this.cbComposit.Name = "cbComposit";
            this.cbComposit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbComposit_Click);
            // 
            // cbQplot
            // 
            this.cbQplot.Label = "q plot";
            this.cbQplot.Name = "cbQplot";
            this.cbQplot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbQplot_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // tglTaskPane
            // 
            this.tglTaskPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tglTaskPane.Image = global::GINtool.Properties.Resources.check;
            this.tglTaskPane.Label = "show/hide steps";
            this.tglTaskPane.Name = "tglTaskPane";
            this.tglTaskPane.ShowImage = true;
            this.tglTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglTaskPane_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.splitButton1);
            this.group3.Items.Add(this.splitButton2);
            this.group3.Label = "reference files";
            this.group3.Name = "group3";
            // 
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = global::GINtool.Properties.Resources.stack1;
            this.splitButton1.Items.Add(this.btnSelectRegulonFile);
            this.splitButton1.Items.Add(this.separator3);
            this.splitButton1.Items.Add(this.btnRegulonFileName);
            this.splitButton1.Label = "regulon file";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.ScreenTip = "select the reference file that contains the mapping tables";
            // 
            // btnSelectRegulonFile
            // 
            this.btnSelectRegulonFile.Image = global::GINtool.Properties.Resources.cursor;
            this.btnSelectRegulonFile.Label = "select";
            this.btnSelectRegulonFile.Name = "btnSelectRegulonFile";
            this.btnSelectRegulonFile.ScreenTip = "select the file";
            this.btnSelectRegulonFile.ShowImage = true;
            this.btnSelectRegulonFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // btnRegulonFileName
            // 
            this.btnRegulonFileName.Enabled = false;
            this.btnRegulonFileName.Label = "no file selected";
            this.btnRegulonFileName.Name = "btnRegulonFileName";
            this.btnRegulonFileName.ShowImage = true;
            // 
            // splitButton2
            // 
            this.splitButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton2.Image = global::GINtool.Properties.Resources.stack2;
            this.splitButton2.Items.Add(this.btnSelectOperonFile);
            this.splitButton2.Items.Add(this.btnResetOperonFile);
            this.splitButton2.Items.Add(this.separator4);
            this.splitButton2.Items.Add(this.btnOperonFile);
            this.splitButton2.Label = "operon file";
            this.splitButton2.Name = "splitButton2";
            this.splitButton2.ScreenTip = "select the reference file that contains the operon-gene mapping";
            // 
            // btnSelectOperonFile
            // 
            this.btnSelectOperonFile.Image = global::GINtool.Properties.Resources.cursor;
            this.btnSelectOperonFile.Label = "select";
            this.btnSelectOperonFile.Name = "btnSelectOperonFile";
            this.btnSelectOperonFile.ScreenTip = "select the file";
            this.btnSelectOperonFile.ShowImage = true;
            this.btnSelectOperonFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectOperonFile_Click);
            // 
            // btnResetOperonFile
            // 
            this.btnResetOperonFile.Image = global::GINtool.Properties.Resources.denied;
            this.btnResetOperonFile.Label = "clear";
            this.btnResetOperonFile.Name = "btnResetOperonFile";
            this.btnResetOperonFile.ShowImage = true;
            this.btnResetOperonFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetOperonFile_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // btnOperonFile
            // 
            this.btnOperonFile.Enabled = false;
            this.btnOperonFile.Label = "no file selected";
            this.btnOperonFile.Name = "btnOperonFile";
            this.btnOperonFile.ShowImage = true;
            // 
            // group6
            // 
            this.group6.Items.Add(this.ddBSU);
            this.group6.Items.Add(this.ddRegulon);
            this.group6.Items.Add(this.ddGene);
            this.group6.Label = "column mappings";
            this.group6.Name = "group6";
            // 
            // ddBSU
            // 
            this.ddBSU.Enabled = false;
            this.ddBSU.Image = global::GINtool.Properties.Resources.target;
            this.ddBSU.Label = "bsu";
            this.ddBSU.Name = "ddBSU";
            this.ddBSU.ScreenTip = "specify the column that contains the bsu code";
            this.ddBSU.ShowImage = true;
            this.ddBSU.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddBSU_SelectionChanged);
            // 
            // ddRegulon
            // 
            this.ddRegulon.Enabled = false;
            this.ddRegulon.Image = global::GINtool.Properties.Resources.crossroads;
            this.ddRegulon.Label = "regulon";
            this.ddRegulon.Name = "ddRegulon";
            this.ddRegulon.ScreenTip = "specify the column that contains the regulator names";
            this.ddRegulon.ShowImage = true;
            this.ddRegulon.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddRegulon_SelectionChanged);
            // 
            // ddGene
            // 
            this.ddGene.Image = global::GINtool.Properties.Resources.dna;
            this.ddGene.Label = "gene";
            this.ddGene.Name = "ddGene";
            this.ddGene.ScreenTip = "specify the column that contains the gene names";
            this.ddGene.ShowImage = true;
            this.ddGene.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddGene_SelectionChanged);
            // 
            // group5
            // 
            this.group5.Items.Add(this.ddDir);
            this.group5.Items.Add(this.btRegDirMap);
            this.group5.Label = "up/down definition";
            this.group5.Name = "group5";
            // 
            // ddDir
            // 
            this.ddDir.Image = global::GINtool.Properties.Resources.traffic;
            this.ddDir.Label = "direction column";
            this.ddDir.Name = "ddDir";
            this.ddDir.ScreenTip = "specify the column that contains the direction definitions";
            this.ddDir.ShowImage = true;
            this.ddDir.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddDir_SelectionChanged);
            // 
            // btRegDirMap
            // 
            this.btRegDirMap.Enabled = false;
            this.btRegDirMap.Image = global::GINtool.Properties.Resources.settings;
            this.btRegDirMap.Label = "regulon direction map";
            this.btRegDirMap.Name = "btRegDirMap";
            this.btRegDirMap.ScreenTip = "define the text mappings for the direction of the regulons";
            this.btRegDirMap.ShowImage = true;
            this.btRegDirMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btRegDirMap_Click);
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
            this.ebLow.ScreenTip = "Set the value for the minimum FC category";
            this.ebLow.Text = null;
            this.ebLow.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ebLow_TextChanged);
            // 
            // ebMid
            // 
            this.ebMid.Label = "mid";
            this.ebMid.Name = "ebMid";
            this.ebMid.ScreenTip = "Set the value for the medium FC category";
            this.ebMid.Text = null;
            this.ebMid.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ebMid_TextChanged);
            // 
            // ebHigh
            // 
            this.ebHigh.Label = "high";
            this.ebHigh.Name = "ebHigh";
            this.ebHigh.ScreenTip = "Set the value for the highest FC category";
            this.ebHigh.Text = null;
            this.ebHigh.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ebHigh_TextChanged);
            // 
            // group4
            // 
            this.group4.Items.Add(this.editMinPval);
            this.group4.Label = "cut-offs";
            this.group4.Name = "group4";
            // 
            // editMinPval
            // 
            this.editMinPval.Label = "p-value";
            this.editMinPval.Name = "editMinPval";
            this.editMinPval.ScreenTip = "Define the p-value cut-off value to include in the comined report";
            this.editMinPval.Text = null;
            this.editMinPval.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editMinPval_TextChanged);
            // 
            // grpPref
            // 
            this.grpPref.Items.Add(this.btnPalette);
            this.grpPref.Items.Add(this.edtMaxGroups);
            this.grpPref.Label = "preferences";
            this.grpPref.Name = "grpPref";
            // 
            // btnPalette
            // 
            this.btnPalette.Buttons.Add(this.clrBerry);
            this.btnPalette.Buttons.Add(this.clrBright);
            this.btnPalette.Buttons.Add(this.clrBrightPastel);
            this.btnPalette.Buttons.Add(this.clrChocolate);
            this.btnPalette.Buttons.Add(this.clrExcel);
            this.btnPalette.Buttons.Add(this.clrGray);
            this.btnPalette.Image = global::GINtool.Properties.Resources.BWJrU_0005_Excel;
            this.btnPalette.Label = "color palette";
            this.btnPalette.Name = "btnPalette";
            this.btnPalette.ShowImage = true;
            // 
            // clrBerry
            // 
            this.clrBerry.Image = global::GINtool.Properties.Resources.BWJrU_0000_Berry;
            this.clrBerry.Label = "Berry";
            this.clrBerry.Name = "clrBerry";
            this.clrBerry.ShowImage = true;
            this.clrBerry.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clrBerry_Click);
            // 
            // clrBright
            // 
            this.clrBright.Image = global::GINtool.Properties.Resources.BWJrU_0001_Bright;
            this.clrBright.Label = "Bright";
            this.clrBright.Name = "clrBright";
            this.clrBright.ShowImage = true;
            this.clrBright.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clrBright_Click);
            // 
            // clrBrightPastel
            // 
            this.clrBrightPastel.Image = global::GINtool.Properties.Resources.BWJrU_0002_Bright_Pastel;
            this.clrBrightPastel.Label = "Bright Pastel";
            this.clrBrightPastel.Name = "clrBrightPastel";
            this.clrBrightPastel.ShowImage = true;
            this.clrBrightPastel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clrBrightPastel_Click);
            // 
            // clrChocolate
            // 
            this.clrChocolate.Image = global::GINtool.Properties.Resources.BWJrU_0003_Chocolate;
            this.clrChocolate.Label = "Chocolate";
            this.clrChocolate.Name = "clrChocolate";
            this.clrChocolate.ShowImage = true;
            this.clrChocolate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clrChocolate_Click);
            // 
            // clrExcel
            // 
            this.clrExcel.Image = global::GINtool.Properties.Resources.BWJrU_0005_Excel;
            this.clrExcel.Label = "Excel";
            this.clrExcel.Name = "clrExcel";
            this.clrExcel.ShowImage = true;
            this.clrExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clrExcel_Click);
            // 
            // clrGray
            // 
            this.clrGray.Image = global::GINtool.Properties.Resources.BWJrU_0007_Gray;
            this.clrGray.Label = "Gray";
            this.clrGray.Name = "clrGray";
            this.clrGray.ShowImage = true;
            this.clrGray.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clrGray_Click_1);
            // 
            // edtMaxGroups
            // 
            this.edtMaxGroups.Label = "max groups";
            this.edtMaxGroups.Name = "edtMaxGroups";
            this.edtMaxGroups.Text = null;
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
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.grpPref.ResumeLayout(false);
            this.grpPref.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabGINtool;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddRegulon;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddBSU;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btRegDirMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddDir;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebLow;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebMid;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebHigh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btApply;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectRegulonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRegulonFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddGene;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editMinPval;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton but_pvalues;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton but_fc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tglTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitbtnEA;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbComposit;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbQplot;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPref;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery btnPalette;
        private Microsoft.Office.Tools.Ribbon.RibbonButton clrExcel;
        private Microsoft.Office.Tools.Ribbon.RibbonButton clrBerry;
        private Microsoft.Office.Tools.Ribbon.RibbonButton clrBright;
        private Microsoft.Office.Tools.Ribbon.RibbonButton clrBrightPastel;
        private Microsoft.Office.Tools.Ribbon.RibbonButton clrChocolate;
        private Microsoft.Office.Tools.Ribbon.RibbonButton clrGray;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtMaxGroups;
    }

    partial class ThisRibbonCollection
    {
        internal GinRibbon GinRibbon
        {
            get { return this.GetRibbon<GinRibbon>(); }
        }
    }
}
