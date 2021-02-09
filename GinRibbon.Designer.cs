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
            this.btLoad = this.Factory.CreateRibbonButton();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.tglTaskPane = this.Factory.CreateRibbonToggleButton();
            this.grpDta = this.Factory.CreateRibbonGroup();
            this.btnSelect = this.Factory.CreateRibbonButton();
            this.grpTable = this.Factory.CreateRibbonGroup();
            this.cbUsePValues = this.Factory.CreateRibbonCheckBox();
            this.cbUseFoldChanges = this.Factory.CreateRibbonCheckBox();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.cbMapping = this.Factory.CreateRibbonCheckBox();
            this.cbSummary = this.Factory.CreateRibbonCheckBox();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.cbCombined = this.Factory.CreateRibbonCheckBox();
            this.cbOperon = this.Factory.CreateRibbonCheckBox();
            this.btApply = this.Factory.CreateRibbonButton();
            this.grpPlot = this.Factory.CreateRibbonGroup();
            this.cbOrderFC = this.Factory.CreateRibbonCheckBox();
            this.cbUseCategories = this.Factory.CreateRibbonCheckBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.cbDistribution = this.Factory.CreateRibbonCheckBox();
            this.cbClustered = this.Factory.CreateRibbonCheckBox();
            this.chkRegulon = this.Factory.CreateRibbonCheckBox();
            this.btPlot = this.Factory.CreateRibbonButton();
            this.grpReference = this.Factory.CreateRibbonGroup();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectRegulonFile = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnRegulonFileName = this.Factory.CreateRibbonButton();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectOperonFile = this.Factory.CreateRibbonButton();
            this.btnResetOperonFile = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.btnOperonFile = this.Factory.CreateRibbonButton();
            this.splitButton4 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectCatFile = this.Factory.CreateRibbonButton();
            this.btnClearCatFile = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnCatFile = this.Factory.CreateRibbonButton();
            this.grpMap = this.Factory.CreateRibbonGroup();
            this.ddBSU = this.Factory.CreateRibbonDropDown();
            this.ddRegulon = this.Factory.CreateRibbonDropDown();
            this.ddGene = this.Factory.CreateRibbonDropDown();
            this.grpUpDown = this.Factory.CreateRibbonGroup();
            this.ddDir = this.Factory.CreateRibbonDropDown();
            this.btRegDirMap = this.Factory.CreateRibbonButton();
            this.grpFC = this.Factory.CreateRibbonGroup();
            this.ebLow = this.Factory.CreateRibbonEditBox();
            this.ebMid = this.Factory.CreateRibbonEditBox();
            this.ebHigh = this.Factory.CreateRibbonEditBox();
            this.grpCutOff = this.Factory.CreateRibbonGroup();
            this.editMinPval = this.Factory.CreateRibbonEditBox();
            this.grpDirection = this.Factory.CreateRibbonGroup();
            this.cbAscending = this.Factory.CreateRibbonCheckBox();
            this.cbDescending = this.Factory.CreateRibbonCheckBox();
            this.TabGINtool.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpDta.SuspendLayout();
            this.grpTable.SuspendLayout();
            this.grpPlot.SuspendLayout();
            this.grpReference.SuspendLayout();
            this.grpMap.SuspendLayout();
            this.grpUpDown.SuspendLayout();
            this.grpFC.SuspendLayout();
            this.grpCutOff.SuspendLayout();
            this.grpDirection.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabGINtool
            // 
            this.TabGINtool.Groups.Add(this.group1);
            this.TabGINtool.Groups.Add(this.grpDta);
            this.TabGINtool.Groups.Add(this.grpTable);
            this.TabGINtool.Groups.Add(this.grpPlot);
            this.TabGINtool.Groups.Add(this.grpReference);
            this.TabGINtool.Groups.Add(this.grpMap);
            this.TabGINtool.Groups.Add(this.grpUpDown);
            this.TabGINtool.Groups.Add(this.grpFC);
            this.TabGINtool.Groups.Add(this.grpCutOff);
            this.TabGINtool.Groups.Add(this.grpDirection);
            this.TabGINtool.Label = "GIN tool";
            this.TabGINtool.Name = "TabGINtool";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btLoad);
            this.group1.Items.Add(this.toggleButton1);
            this.group1.Items.Add(this.tglTaskPane);
            this.group1.Label = "main";
            this.group1.Name = "group1";
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
            // toggleButton1
            // 
            this.toggleButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton1.Image = global::GINtool.Properties.Resources.tools;
            this.toggleButton1.Label = "show/hide settings";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButton1_Click);
            // 
            // tglTaskPane
            // 
            this.tglTaskPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tglTaskPane.Image = global::GINtool.Properties.Resources.clipboard;
            this.tglTaskPane.Label = "show/hide manual";
            this.tglTaskPane.Name = "tglTaskPane";
            this.tglTaskPane.ShowImage = true;
            this.tglTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglTaskPane_Click);
            // 
            // grpDta
            // 
            this.grpDta.Items.Add(this.btnSelect);
            this.grpDta.Label = "data";
            this.grpDta.Name = "grpDta";
            // 
            // btnSelect
            // 
            this.btnSelect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSelect.Image = global::GINtool.Properties.Resources.crop;
            this.btnSelect.Label = "select";
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.ShowImage = true;
            this.btnSelect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelect_Click);
            // 
            // grpTable
            // 
            this.grpTable.Items.Add(this.cbUsePValues);
            this.grpTable.Items.Add(this.cbUseFoldChanges);
            this.grpTable.Items.Add(this.separator6);
            this.grpTable.Items.Add(this.cbMapping);
            this.grpTable.Items.Add(this.cbSummary);
            this.grpTable.Items.Add(this.separator5);
            this.grpTable.Items.Add(this.cbCombined);
            this.grpTable.Items.Add(this.cbOperon);
            this.grpTable.Items.Add(this.btApply);
            this.grpTable.Label = "tables";
            this.grpTable.Name = "grpTable";
            // 
            // cbUsePValues
            // 
            this.cbUsePValues.Label = "p-values";
            this.cbUsePValues.Name = "cbUsePValues";
            this.cbUsePValues.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbUsePValues_Click);
            // 
            // cbUseFoldChanges
            // 
            this.cbUseFoldChanges.Label = "fold-changes";
            this.cbUseFoldChanges.Name = "cbUseFoldChanges";
            this.cbUseFoldChanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbUseFoldChanges_Click);
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // cbMapping
            // 
            this.cbMapping.Checked = true;
            this.cbMapping.Label = "mapping";
            this.cbMapping.Name = "cbMapping";
            this.cbMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbMapping_Click);
            // 
            // cbSummary
            // 
            this.cbSummary.Checked = true;
            this.cbSummary.Label = "summary";
            this.cbSummary.Name = "cbSummary";
            this.cbSummary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbSummary_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // cbCombined
            // 
            this.cbCombined.Checked = true;
            this.cbCombined.Label = "combined";
            this.cbCombined.Name = "cbCombined";
            this.cbCombined.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbCombined_Click);
            // 
            // cbOperon
            // 
            this.cbOperon.Label = "operon";
            this.cbOperon.Name = "cbOperon";
            this.cbOperon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbOperon_Click);
            // 
            // btApply
            // 
            this.btApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btApply.Image = global::GINtool.Properties.Resources.check1;
            this.btApply.Label = "make tables";
            this.btApply.Name = "btApply";
            this.btApply.ScreenTip = "Start the analysis (PVALUE,FC,BSU)";
            this.btApply.ShowImage = true;
            this.btApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btApply_Click);
            // 
            // grpPlot
            // 
            this.grpPlot.Items.Add(this.cbOrderFC);
            this.grpPlot.Items.Add(this.cbUseCategories);
            this.grpPlot.Items.Add(this.separator2);
            this.grpPlot.Items.Add(this.cbDistribution);
            this.grpPlot.Items.Add(this.cbClustered);
            this.grpPlot.Items.Add(this.chkRegulon);
            this.grpPlot.Items.Add(this.btPlot);
            this.grpPlot.Label = "plots";
            this.grpPlot.Name = "grpPlot";
            // 
            // cbOrderFC
            // 
            this.cbOrderFC.Label = "sort results";
            this.cbOrderFC.Name = "cbOrderFC";
            this.cbOrderFC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbOrderFC_Click);
            // 
            // cbUseCategories
            // 
            this.cbUseCategories.Label = "use categories";
            this.cbUseCategories.Name = "cbUseCategories";
            this.cbUseCategories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbUseCategories_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // cbDistribution
            // 
            this.cbDistribution.Label = "distribution";
            this.cbDistribution.Name = "cbDistribution";
            this.cbDistribution.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbDistribution_Click);
            // 
            // cbClustered
            // 
            this.cbClustered.Label = "category/regulon";
            this.cbClustered.Name = "cbClustered";
            this.cbClustered.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbClustered_Click);
            // 
            // chkRegulon
            // 
            this.chkRegulon.Label = "regulon mean/var";
            this.chkRegulon.Name = "chkRegulon";
            this.chkRegulon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
            // 
            // btPlot
            // 
            this.btPlot.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btPlot.Image = global::GINtool.Properties.Resources.barchart;
            this.btPlot.Label = "make plots";
            this.btPlot.Name = "btPlot";
            this.btPlot.ShowImage = true;
            this.btPlot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btPlot_Click);
            // 
            // grpReference
            // 
            this.grpReference.Items.Add(this.splitButton1);
            this.grpReference.Items.Add(this.splitButton2);
            this.grpReference.Items.Add(this.splitButton4);
            this.grpReference.Label = "reference files";
            this.grpReference.Name = "grpReference";
            this.grpReference.Visible = false;
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
            // splitButton4
            // 
            this.splitButton4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton4.Image = global::GINtool.Properties.Resources.swatches;
            this.splitButton4.Items.Add(this.btnSelectCatFile);
            this.splitButton4.Items.Add(this.btnClearCatFile);
            this.splitButton4.Items.Add(this.separator1);
            this.splitButton4.Items.Add(this.btnCatFile);
            this.splitButton4.Label = "category file";
            this.splitButton4.Name = "splitButton4";
            // 
            // btnSelectCatFile
            // 
            this.btnSelectCatFile.Image = global::GINtool.Properties.Resources.cursor;
            this.btnSelectCatFile.Label = "select";
            this.btnSelectCatFile.Name = "btnSelectCatFile";
            this.btnSelectCatFile.ShowImage = true;
            this.btnSelectCatFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectCatFile_Click);
            // 
            // btnClearCatFile
            // 
            this.btnClearCatFile.Image = global::GINtool.Properties.Resources.denied;
            this.btnClearCatFile.Label = "clear";
            this.btnClearCatFile.Name = "btnClearCatFile";
            this.btnClearCatFile.ShowImage = true;
            this.btnClearCatFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearCatFile_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnCatFile
            // 
            this.btnCatFile.Label = "no file selected";
            this.btnCatFile.Name = "btnCatFile";
            this.btnCatFile.ShowImage = true;
            // 
            // grpMap
            // 
            this.grpMap.Items.Add(this.ddBSU);
            this.grpMap.Items.Add(this.ddRegulon);
            this.grpMap.Items.Add(this.ddGene);
            this.grpMap.Label = "column mappings";
            this.grpMap.Name = "grpMap";
            this.grpMap.Visible = false;
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
            // grpUpDown
            // 
            this.grpUpDown.Items.Add(this.ddDir);
            this.grpUpDown.Items.Add(this.btRegDirMap);
            this.grpUpDown.Label = "up/down definition";
            this.grpUpDown.Name = "grpUpDown";
            this.grpUpDown.Visible = false;
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
            this.btRegDirMap.Label = "regulon direction mapping";
            this.btRegDirMap.Name = "btRegDirMap";
            this.btRegDirMap.ScreenTip = "define the text mappings for the direction of the regulons";
            this.btRegDirMap.ShowImage = true;
            this.btRegDirMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btRegDirMap_Click);
            // 
            // grpFC
            // 
            this.grpFC.Items.Add(this.ebLow);
            this.grpFC.Items.Add(this.ebMid);
            this.grpFC.Items.Add(this.ebHigh);
            this.grpFC.Label = "fc ranges";
            this.grpFC.Name = "grpFC";
            this.grpFC.Visible = false;
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
            // grpCutOff
            // 
            this.grpCutOff.Items.Add(this.editMinPval);
            this.grpCutOff.Label = "cut-offs";
            this.grpCutOff.Name = "grpCutOff";
            this.grpCutOff.Visible = false;
            // 
            // editMinPval
            // 
            this.editMinPval.Label = "p-value";
            this.editMinPval.Name = "editMinPval";
            this.editMinPval.ScreenTip = "Define the p-value cut-off value to include in the comined report";
            this.editMinPval.ShowImage = true;
            this.editMinPval.Text = null;
            this.editMinPval.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editMinPval_TextChanged);
            // 
            // grpDirection
            // 
            this.grpDirection.Items.Add(this.cbAscending);
            this.grpDirection.Items.Add(this.cbDescending);
            this.grpDirection.Label = "sort direction";
            this.grpDirection.Name = "grpDirection";
            this.grpDirection.Visible = false;
            // 
            // cbAscending
            // 
            this.cbAscending.Label = "ascending";
            this.cbAscending.Name = "cbAscending";
            this.cbAscending.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbAscending_Click);
            // 
            // cbDescending
            // 
            this.cbDescending.Label = "descending";
            this.cbDescending.Name = "cbDescending";
            this.cbDescending.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbDescending_Click);
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
            this.grpDta.ResumeLayout(false);
            this.grpDta.PerformLayout();
            this.grpTable.ResumeLayout(false);
            this.grpTable.PerformLayout();
            this.grpPlot.ResumeLayout(false);
            this.grpPlot.PerformLayout();
            this.grpReference.ResumeLayout(false);
            this.grpReference.PerformLayout();
            this.grpMap.ResumeLayout(false);
            this.grpMap.PerformLayout();
            this.grpUpDown.ResumeLayout(false);
            this.grpUpDown.PerformLayout();
            this.grpFC.ResumeLayout(false);
            this.grpFC.PerformLayout();
            this.grpCutOff.ResumeLayout(false);
            this.grpCutOff.PerformLayout();
            this.grpDirection.ResumeLayout(false);
            this.grpDirection.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFC;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebLow;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebMid;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebHigh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btApply;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpReference;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectRegulonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRegulonFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMap;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddGene;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUpDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCutOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editMinPval;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tglTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPlot;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUseCategories;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btPlot;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbOrderFC;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectCatFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearCatFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCatFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDistribution;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbClustered;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkRegulon;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbCombined;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbSummary;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbOperon;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUsePValues;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUseFoldChanges;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDta;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDirection;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAscending;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDescending;
    }

    partial class ThisRibbonCollection
    {
        internal GinRibbon GinRibbon
        {
            get { return this.GetRibbon<GinRibbon>(); }
        }
    }
}
