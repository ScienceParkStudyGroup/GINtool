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
            this.grpDta = this.Factory.CreateRibbonGroup();
            this.btnSelect = this.Factory.CreateRibbonButton();
            this.grpFilter = this.Factory.CreateRibbonGroup();
            this.cbNoFilter = this.Factory.CreateRibbonCheckBox();
            this.cbUsePValues = this.Factory.CreateRibbonCheckBox();
            this.cbUseFoldChanges = this.Factory.CreateRibbonCheckBox();
            this.grpFocus = this.Factory.CreateRibbonGroup();
            this.cbUseOperons = this.Factory.CreateRibbonCheckBox();
            this.cbUseCategories = this.Factory.CreateRibbonCheckBox();
            this.cbUseRegulons = this.Factory.CreateRibbonCheckBox();
            this.grpTable = this.Factory.CreateRibbonGroup();
            this.cbMapping = this.Factory.CreateRibbonCheckBox();
            this.cbSummary = this.Factory.CreateRibbonCheckBox();
            this.dbTblRanking = this.Factory.CreateRibbonCheckBox();
            this.cbCombined = this.Factory.CreateRibbonCheckBox();
            this.btApply = this.Factory.CreateRibbonButton();
            this.grpPlot = this.Factory.CreateRibbonGroup();
            this.cbDistribution = this.Factory.CreateRibbonCheckBox();
            this.cbClustered = this.Factory.CreateRibbonCheckBox();
            this.chkRegulon = this.Factory.CreateRibbonCheckBox();
            this.btPlot = this.Factory.CreateRibbonButton();
            this.grpReference = this.Factory.CreateRibbonGroup();
            this.splitBtnGenesFile = this.Factory.CreateRibbonSplitButton();
            this.btnSelectGenesFile = this.Factory.CreateRibbonButton();
            this.cbGenesFileMapping = this.Factory.CreateRibbonCheckBox();
            this.btnClearGenInfo = this.Factory.CreateRibbonButton();
            this.separator7 = this.Factory.CreateRibbonSeparator();
            this.btnGenesFileSelected = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectOperonFile = this.Factory.CreateRibbonButton();
            this.btnResetOperonFile = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.btnOperonFile = this.Factory.CreateRibbonButton();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectRegulonFile = this.Factory.CreateRibbonButton();
            this.cbRegulonMapping = this.Factory.CreateRibbonCheckBox();
            this.btnResetRegulonLinkageFile = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnRegulonFileName = this.Factory.CreateRibbonButton();
            this.splitButton4 = this.Factory.CreateRibbonSplitButton();
            this.btnSelectCatFile = this.Factory.CreateRibbonButton();
            this.cbCategoryMapping = this.Factory.CreateRibbonCheckBox();
            this.btnClearCatFile = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnCatFile = this.Factory.CreateRibbonButton();
            this.separator9 = this.Factory.CreateRibbonSeparator();
            this.splitButton3 = this.Factory.CreateRibbonSplitButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.cbRegInfoColumnMapping = this.Factory.CreateRibbonCheckBox();
            this.btnClearRegulonInfo = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnRegInfoFileName = this.Factory.CreateRibbonButton();
            this.grpGenesMapping = this.Factory.CreateRibbonGroup();
            this.ddGnsName = this.Factory.CreateRibbonDropDown();
            this.ddGenesBSU = this.Factory.CreateRibbonDropDown();
            this.ddGenesFunction = this.Factory.CreateRibbonDropDown();
            this.separator8 = this.Factory.CreateRibbonSeparator();
            this.ddGenesDescription = this.Factory.CreateRibbonDropDown();
            this.grpMap = this.Factory.CreateRibbonGroup();
            this.ddGene = this.Factory.CreateRibbonDropDown();
            this.ddBSU = this.Factory.CreateRibbonDropDown();
            this.ddRegulon = this.Factory.CreateRibbonDropDown();
            this.separator10 = this.Factory.CreateRibbonSeparator();
            this.ddDir = this.Factory.CreateRibbonDropDown();
            this.btRegDirMap = this.Factory.CreateRibbonButton();
            this.grpRegulonInfo = this.Factory.CreateRibbonGroup();
            this.ddRegInfoId = this.Factory.CreateRibbonDropDown();
            this.ddRegInfoSize = this.Factory.CreateRibbonDropDown();
            this.ddRegInfoFunction = this.Factory.CreateRibbonDropDown();
            this.grpColMapCategory = this.Factory.CreateRibbonGroup();
            this.ddCatID = this.Factory.CreateRibbonDropDown();
            this.ddCatName = this.Factory.CreateRibbonDropDown();
            this.ddCatBSU = this.Factory.CreateRibbonDropDown();
            this.grpCutOff = this.Factory.CreateRibbonGroup();
            this.editMinPval = this.Factory.CreateRibbonEditBox();
            this.ebLow = this.Factory.CreateRibbonEditBox();
            this.grpDirection = this.Factory.CreateRibbonGroup();
            this.cbAscending = this.Factory.CreateRibbonCheckBox();
            this.cbDescending = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.tglTaskPane = this.Factory.CreateRibbonToggleButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.TabGINtool.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpDta.SuspendLayout();
            this.grpFilter.SuspendLayout();
            this.grpFocus.SuspendLayout();
            this.grpTable.SuspendLayout();
            this.grpPlot.SuspendLayout();
            this.grpReference.SuspendLayout();
            this.grpGenesMapping.SuspendLayout();
            this.grpMap.SuspendLayout();
            this.grpRegulonInfo.SuspendLayout();
            this.grpColMapCategory.SuspendLayout();
            this.grpCutOff.SuspendLayout();
            this.grpDirection.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabGINtool
            // 
            this.TabGINtool.Groups.Add(this.group1);
            this.TabGINtool.Groups.Add(this.grpDta);
            this.TabGINtool.Groups.Add(this.grpFilter);
            this.TabGINtool.Groups.Add(this.grpFocus);
            this.TabGINtool.Groups.Add(this.grpTable);
            this.TabGINtool.Groups.Add(this.grpPlot);
            this.TabGINtool.Groups.Add(this.grpReference);
            this.TabGINtool.Groups.Add(this.grpGenesMapping);
            this.TabGINtool.Groups.Add(this.grpMap);
            this.TabGINtool.Groups.Add(this.grpRegulonInfo);
            this.TabGINtool.Groups.Add(this.grpColMapCategory);
            this.TabGINtool.Groups.Add(this.grpCutOff);
            this.TabGINtool.Groups.Add(this.grpDirection);
            this.TabGINtool.Groups.Add(this.group2);
            this.TabGINtool.Label = "GIN tool";
            this.TabGINtool.Name = "TabGINtool";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btLoad);
            this.group1.Items.Add(this.toggleButton1);
            this.group1.Label = "main";
            this.group1.Name = "group1";
            // 
            // btLoad
            // 
            this.btLoad.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btLoad.Enabled = false;
            this.btLoad.Image = global::GINtool.Properties.Resources.stack;
            this.btLoad.Label = "initialize data";
            this.btLoad.Name = "btLoad";
            this.btLoad.ScreenTip = "load reference data into memory";
            this.btLoad.ShowImage = true;
            this.btLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_Load_Click);
            // 
            // toggleButton1
            // 
            this.toggleButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButton1.Image = global::GINtool.Properties.Resources.tools;
            this.toggleButton1.Label = "show/hide settings";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            this.toggleButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Toggle_Settings_Click);
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
            this.btnSelect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_Select_Click);
            // 
            // grpFilter
            // 
            this.grpFilter.Items.Add(this.cbNoFilter);
            this.grpFilter.Items.Add(this.cbUsePValues);
            this.grpFilter.Items.Add(this.cbUseFoldChanges);
            this.grpFilter.Label = "filter settings";
            this.grpFilter.Name = "grpFilter";
            // 
            // cbNoFilter
            // 
            this.cbNoFilter.Label = "no filter";
            this.cbNoFilter.Name = "cbNoFilter";
            this.cbNoFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbNoFilter_Click);
            // 
            // cbUsePValues
            // 
            this.cbUsePValues.Label = "p-values";
            this.cbUsePValues.Name = "cbUsePValues";
            this.cbUsePValues.ScreenTip = "select items to include based on p-value cutoff settings";
            this.cbUsePValues.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_UsePValues_Click);
            // 
            // cbUseFoldChanges
            // 
            this.cbUseFoldChanges.Label = "fold-changes";
            this.cbUseFoldChanges.Name = "cbUseFoldChanges";
            this.cbUseFoldChanges.ScreenTip = "select items to include based on fold-changes settings";
            this.cbUseFoldChanges.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_UseFoldChanges_Click);
            // 
            // grpFocus
            // 
            this.grpFocus.Items.Add(this.cbUseOperons);
            this.grpFocus.Items.Add(this.cbUseCategories);
            this.grpFocus.Items.Add(this.cbUseRegulons);
            this.grpFocus.Label = "focus";
            this.grpFocus.Name = "grpFocus";
            // 
            // cbUseOperons
            // 
            this.cbUseOperons.Label = "use operons";
            this.cbUseOperons.Name = "cbUseOperons";
            this.cbUseOperons.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Operon_Click);
            // 
            // cbUseCategories
            // 
            this.cbUseCategories.Label = "use categories";
            this.cbUseCategories.Name = "cbUseCategories";
            this.cbUseCategories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_UseCategories_Click);
            // 
            // cbUseRegulons
            // 
            this.cbUseRegulons.Label = "use regulons";
            this.cbUseRegulons.Name = "cbUseRegulons";
            this.cbUseRegulons.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_UseRegulons_Click);
            // 
            // grpTable
            // 
            this.grpTable.Items.Add(this.cbMapping);
            this.grpTable.Items.Add(this.cbSummary);
            this.grpTable.Items.Add(this.dbTblRanking);
            this.grpTable.Items.Add(this.cbCombined);
            this.grpTable.Items.Add(this.btApply);
            this.grpTable.Label = "tables";
            this.grpTable.Name = "grpTable";
            // 
            // cbMapping
            // 
            this.cbMapping.Checked = true;
            this.cbMapping.Label = "details";
            this.cbMapping.Name = "cbMapping";
            this.cbMapping.Visible = false;
            this.cbMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Mapping_Click);
            // 
            // cbSummary
            // 
            this.cbSummary.Label = "summary";
            this.cbSummary.Name = "cbSummary";
            this.cbSummary.Visible = false;
            this.cbSummary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Summary_Click);
            // 
            // dbTblRanking
            // 
            this.dbTblRanking.Checked = true;
            this.dbTblRanking.Enabled = false;
            this.dbTblRanking.Label = "ranking";
            this.dbTblRanking.Name = "dbTblRanking";
            this.dbTblRanking.ScreenTip = "create a table with ranked regulons/categories";
            this.dbTblRanking.Visible = false;
            // 
            // cbCombined
            // 
            this.cbCombined.Checked = true;
            this.cbCombined.Label = "combined";
            this.cbCombined.Name = "cbCombined";
            this.cbCombined.Visible = false;
            this.cbCombined.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Combined_Click);
            // 
            // btApply
            // 
            this.btApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btApply.Image = global::GINtool.Properties.Resources.check1;
            this.btApply.Label = "make tables";
            this.btApply.Name = "btApply";
            this.btApply.ScreenTip = "Start the analysis (PVALUE,FC,BSU)";
            this.btApply.ShowImage = true;
            this.btApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_Apply_Click);
            // 
            // grpPlot
            // 
            this.grpPlot.Items.Add(this.cbDistribution);
            this.grpPlot.Items.Add(this.cbClustered);
            this.grpPlot.Items.Add(this.chkRegulon);
            this.grpPlot.Items.Add(this.btPlot);
            this.grpPlot.Label = "plots";
            this.grpPlot.Name = "grpPlot";
            // 
            // cbDistribution
            // 
            this.cbDistribution.Label = "distribution";
            this.cbDistribution.Name = "cbDistribution";
            this.cbDistribution.Visible = false;
            this.cbDistribution.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Distribution_Click);
            // 
            // cbClustered
            // 
            this.cbClustered.Label = "spreading";
            this.cbClustered.Name = "cbClustered";
            this.cbClustered.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Spreading_Click);
            // 
            // chkRegulon
            // 
            this.chkRegulon.Label = "ranking";
            this.chkRegulon.Name = "chkRegulon";
            this.chkRegulon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Regulon_Click);
            // 
            // btPlot
            // 
            this.btPlot.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btPlot.Image = global::GINtool.Properties.Resources.barchart;
            this.btPlot.Label = "make plots";
            this.btPlot.Name = "btPlot";
            this.btPlot.ShowImage = true;
            this.btPlot.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_Plot_Click);
            // 
            // grpReference
            // 
            this.grpReference.Items.Add(this.splitBtnGenesFile);
            this.grpReference.Items.Add(this.separator5);
            this.grpReference.Items.Add(this.splitButton2);
            this.grpReference.Items.Add(this.splitButton1);
            this.grpReference.Items.Add(this.splitButton4);
            this.grpReference.Items.Add(this.separator9);
            this.grpReference.Items.Add(this.splitButton3);
            this.grpReference.Label = "reference files";
            this.grpReference.Name = "grpReference";
            this.grpReference.Visible = false;
            // 
            // splitBtnGenesFile
            // 
            this.splitBtnGenesFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitBtnGenesFile.Image = global::GINtool.Properties.Resources.gene_editing_icon_64x64;
            this.splitBtnGenesFile.Items.Add(this.btnSelectGenesFile);
            this.splitBtnGenesFile.Items.Add(this.cbGenesFileMapping);
            this.splitBtnGenesFile.Items.Add(this.btnClearGenInfo);
            this.splitBtnGenesFile.Items.Add(this.separator7);
            this.splitBtnGenesFile.Items.Add(this.btnGenesFileSelected);
            this.splitBtnGenesFile.Label = "gene info";
            this.splitBtnGenesFile.Name = "splitBtnGenesFile";
            this.splitBtnGenesFile.ScreenTip = "select the csv file that contains the gene description, names and bsu codes";
            this.splitBtnGenesFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitBtnGenesFile_Click);
            // 
            // btnSelectGenesFile
            // 
            this.btnSelectGenesFile.Image = global::GINtool.Properties.Resources.cursor;
            this.btnSelectGenesFile.Label = "select";
            this.btnSelectGenesFile.Name = "btnSelectGenesFile";
            this.btnSelectGenesFile.ShowImage = true;
            this.btnSelectGenesFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectGenesFile_Click);
            // 
            // cbGenesFileMapping
            // 
            this.cbGenesFileMapping.Label = "show/hide column mapping";
            this.cbGenesFileMapping.Name = "cbGenesFileMapping";
            this.cbGenesFileMapping.ScreenTip = "show/hide column mapping for this file";
            this.cbGenesFileMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenesFileMapping_Click);
            // 
            // btnClearGenInfo
            // 
            this.btnClearGenInfo.Image = global::GINtool.Properties.Resources.denied;
            this.btnClearGenInfo.Label = "clear";
            this.btnClearGenInfo.Name = "btnClearGenInfo";
            this.btnClearGenInfo.ShowImage = true;
            this.btnClearGenInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearGenInfo_Click);
            // 
            // separator7
            // 
            this.separator7.Name = "separator7";
            // 
            // btnGenesFileSelected
            // 
            this.btnGenesFileSelected.Enabled = false;
            this.btnGenesFileSelected.Label = "no file selected";
            this.btnGenesFileSelected.Name = "btnGenesFileSelected";
            this.btnGenesFileSelected.ShowImage = true;
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // splitButton2
            // 
            this.splitButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton2.Image = global::GINtool.Properties.Resources.stack2;
            this.splitButton2.Items.Add(this.btnSelectOperonFile);
            this.splitButton2.Items.Add(this.btnResetOperonFile);
            this.splitButton2.Items.Add(this.separator4);
            this.splitButton2.Items.Add(this.btnOperonFile);
            this.splitButton2.Label = "operons";
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
            this.btnSelectOperonFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_SelectOperonFile_Click);
            // 
            // btnResetOperonFile
            // 
            this.btnResetOperonFile.Image = global::GINtool.Properties.Resources.denied;
            this.btnResetOperonFile.Label = "clear";
            this.btnResetOperonFile.Name = "btnResetOperonFile";
            this.btnResetOperonFile.ShowImage = true;
            this.btnResetOperonFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_ResetOperonFile_Click);
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
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = global::GINtool.Properties.Resources.stack1;
            this.splitButton1.Items.Add(this.btnSelectRegulonFile);
            this.splitButton1.Items.Add(this.cbRegulonMapping);
            this.splitButton1.Items.Add(this.btnResetRegulonLinkageFile);
            this.splitButton1.Items.Add(this.separator3);
            this.splitButton1.Items.Add(this.btnRegulonFileName);
            this.splitButton1.Label = "regulons";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.ScreenTip = "select the csv file that contains the regulon mapping info (subtwiki)";
            this.splitButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton1_Click);
            // 
            // btnSelectRegulonFile
            // 
            this.btnSelectRegulonFile.Image = global::GINtool.Properties.Resources.cursor;
            this.btnSelectRegulonFile.Label = "select";
            this.btnSelectRegulonFile.Name = "btnSelectRegulonFile";
            this.btnSelectRegulonFile.ScreenTip = "select the file";
            this.btnSelectRegulonFile.ShowImage = true;
            this.btnSelectRegulonFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonSelectRegulonFile);
            // 
            // cbRegulonMapping
            // 
            this.cbRegulonMapping.Label = "show/hide column mapping";
            this.cbRegulonMapping.Name = "cbRegulonMapping";
            this.cbRegulonMapping.ScreenTip = "show/hide column mappings for this file";
            this.cbRegulonMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
            // 
            // btnResetRegulonLinkageFile
            // 
            this.btnResetRegulonLinkageFile.Image = global::GINtool.Properties.Resources.denied;
            this.btnResetRegulonLinkageFile.Label = "clear";
            this.btnResetRegulonLinkageFile.Name = "btnResetRegulonLinkageFile";
            this.btnResetRegulonLinkageFile.ShowImage = true;
            this.btnResetRegulonLinkageFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetRegulonFile_Click);
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
            // splitButton4
            // 
            this.splitButton4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton4.Image = global::GINtool.Properties.Resources.swatches;
            this.splitButton4.Items.Add(this.btnSelectCatFile);
            this.splitButton4.Items.Add(this.cbCategoryMapping);
            this.splitButton4.Items.Add(this.btnClearCatFile);
            this.splitButton4.Items.Add(this.separator1);
            this.splitButton4.Items.Add(this.btnCatFile);
            this.splitButton4.Label = "categories";
            this.splitButton4.Name = "splitButton4";
            this.splitButton4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton4_Click);
            // 
            // btnSelectCatFile
            // 
            this.btnSelectCatFile.Image = global::GINtool.Properties.Resources.cursor;
            this.btnSelectCatFile.Label = "select";
            this.btnSelectCatFile.Name = "btnSelectCatFile";
            this.btnSelectCatFile.ShowImage = true;
            this.btnSelectCatFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_SelectCatFile_Click);
            // 
            // cbCategoryMapping
            // 
            this.cbCategoryMapping.Label = "show/hide column mapping";
            this.cbCategoryMapping.Name = "cbCategoryMapping";
            this.cbCategoryMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbCategoryMapping_Click);
            // 
            // btnClearCatFile
            // 
            this.btnClearCatFile.Image = global::GINtool.Properties.Resources.denied;
            this.btnClearCatFile.Label = "clear";
            this.btnClearCatFile.Name = "btnClearCatFile";
            this.btnClearCatFile.ShowImage = true;
            this.btnClearCatFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_ClearCatFile_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnCatFile
            // 
            this.btnCatFile.Enabled = false;
            this.btnCatFile.Label = "no file selected";
            this.btnCatFile.Name = "btnCatFile";
            this.btnCatFile.ShowImage = true;
            // 
            // separator9
            // 
            this.separator9.Name = "separator9";
            // 
            // splitButton3
            // 
            this.splitButton3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton3.Image = global::GINtool.Properties.Resources.gene_editing_icon_64x64;
            this.splitButton3.Items.Add(this.button1);
            this.splitButton3.Items.Add(this.cbRegInfoColumnMapping);
            this.splitButton3.Items.Add(this.btnClearRegulonInfo);
            this.splitButton3.Items.Add(this.separator2);
            this.splitButton3.Items.Add(this.btnRegInfoFileName);
            this.splitButton3.Label = "regulon info";
            this.splitButton3.Name = "splitButton3";
            this.splitButton3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton3_Click);
            // 
            // button1
            // 
            this.button1.Image = global::GINtool.Properties.Resources.cursor;
            this.button1.Label = "select";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // cbRegInfoColumnMapping
            // 
            this.cbRegInfoColumnMapping.Label = "show/hide column mapping";
            this.cbRegInfoColumnMapping.Name = "cbRegInfoColumnMapping";
            this.cbRegInfoColumnMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbRegInfoColumnMapping_Click);
            // 
            // btnClearRegulonInfo
            // 
            this.btnClearRegulonInfo.Image = global::GINtool.Properties.Resources.denied;
            this.btnClearRegulonInfo.Label = "clear";
            this.btnClearRegulonInfo.Name = "btnClearRegulonInfo";
            this.btnClearRegulonInfo.ShowImage = true;
            this.btnClearRegulonInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearRegulonInfo_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnRegInfoFileName
            // 
            this.btnRegInfoFileName.Enabled = false;
            this.btnRegInfoFileName.Label = "no file selected";
            this.btnRegInfoFileName.Name = "btnRegInfoFileName";
            this.btnRegInfoFileName.ShowImage = true;
            // 
            // grpGenesMapping
            // 
            this.grpGenesMapping.Items.Add(this.ddGnsName);
            this.grpGenesMapping.Items.Add(this.ddGenesBSU);
            this.grpGenesMapping.Items.Add(this.ddGenesFunction);
            this.grpGenesMapping.Items.Add(this.separator8);
            this.grpGenesMapping.Items.Add(this.ddGenesDescription);
            this.grpGenesMapping.Label = "gene info column mapping";
            this.grpGenesMapping.Name = "grpGenesMapping";
            this.grpGenesMapping.Visible = false;
            // 
            // ddGnsName
            // 
            this.ddGnsName.Image = global::GINtool.Properties.Resources.dna;
            this.ddGnsName.Label = "gene name";
            this.ddGnsName.Name = "ddGnsName";
            this.ddGnsName.ScreenTip = "map the column that contains the genes\' common name";
            this.ddGnsName.ShowImage = true;
            this.ddGnsName.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddGnsName_SelectionChanged);
            // 
            // ddGenesBSU
            // 
            this.ddGenesBSU.Image = global::GINtool.Properties.Resources.target;
            this.ddGenesBSU.Label = "gene id";
            this.ddGenesBSU.Name = "ddGenesBSU";
            this.ddGenesBSU.ScreenTip = "map the column that defines the genes\' bsu code";
            this.ddGenesBSU.ShowImage = true;
            this.ddGenesBSU.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddGenesBSU_SelectionChanged);
            // 
            // ddGenesFunction
            // 
            this.ddGenesFunction.Image = global::GINtool.Properties.Resources.tools;
            this.ddGenesFunction.Label = "function";
            this.ddGenesFunction.Name = "ddGenesFunction";
            this.ddGenesFunction.ScreenTip = "map the column that describes the genes\' function";
            this.ddGenesFunction.ShowImage = true;
            this.ddGenesFunction.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddGenesFunction_SelectionChanged);
            // 
            // separator8
            // 
            this.separator8.Name = "separator8";
            // 
            // ddGenesDescription
            // 
            this.ddGenesDescription.Image = global::GINtool.Properties.Resources.keyboard;
            this.ddGenesDescription.Label = "description";
            this.ddGenesDescription.Name = "ddGenesDescription";
            this.ddGenesDescription.ShowImage = true;
            this.ddGenesDescription.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddGenesDescription_SelectionChanged);
            // 
            // grpMap
            // 
            this.grpMap.Items.Add(this.ddGene);
            this.grpMap.Items.Add(this.ddBSU);
            this.grpMap.Items.Add(this.ddRegulon);
            this.grpMap.Items.Add(this.separator10);
            this.grpMap.Items.Add(this.ddDir);
            this.grpMap.Items.Add(this.btRegDirMap);
            this.grpMap.Label = "regulon linkage column mapping";
            this.grpMap.Name = "grpMap";
            this.grpMap.Visible = false;
            // 
            // ddGene
            // 
            this.ddGene.Image = global::GINtool.Properties.Resources.dna;
            this.ddGene.Label = "gene name";
            this.ddGene.Name = "ddGene";
            this.ddGene.ScreenTip = "specify the column that contains the gene names";
            this.ddGene.ShowImage = true;
            this.ddGene.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown_Gene_SelectionChanged);
            // 
            // ddBSU
            // 
            this.ddBSU.Enabled = false;
            this.ddBSU.Image = global::GINtool.Properties.Resources.target;
            this.ddBSU.Label = "gene id";
            this.ddBSU.Name = "ddBSU";
            this.ddBSU.ScreenTip = "specify the column that contains the gene identifier (BSU) code";
            this.ddBSU.ShowImage = true;
            this.ddBSU.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown_BSU_SelectionChanged);
            // 
            // ddRegulon
            // 
            this.ddRegulon.Enabled = false;
            this.ddRegulon.Image = global::GINtool.Properties.Resources.crossroads;
            this.ddRegulon.Label = "regulon name";
            this.ddRegulon.Name = "ddRegulon";
            this.ddRegulon.ScreenTip = "specify the column that contains the regulon names";
            this.ddRegulon.ShowImage = true;
            this.ddRegulon.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown_Regulon_SelectionChanged);
            // 
            // separator10
            // 
            this.separator10.Name = "separator10";
            // 
            // ddDir
            // 
            this.ddDir.Image = global::GINtool.Properties.Resources.traffic;
            this.ddDir.Label = "direction column";
            this.ddDir.Name = "ddDir";
            this.ddDir.ScreenTip = "specify the column that contains the direction definitions";
            this.ddDir.ShowImage = true;
            this.ddDir.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DropDown_RegulonDirection_SelectionChanged);
            // 
            // btRegDirMap
            // 
            this.btRegDirMap.Enabled = false;
            this.btRegDirMap.Image = global::GINtool.Properties.Resources.settings;
            this.btRegDirMap.Label = "regulon direction";
            this.btRegDirMap.Name = "btRegDirMap";
            this.btRegDirMap.ScreenTip = "define the text mappings for the direction of the regulons";
            this.btRegDirMap.ShowImage = true;
            this.btRegDirMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_RegulonDirectionMap_Click);
            // 
            // grpRegulonInfo
            // 
            this.grpRegulonInfo.Items.Add(this.ddRegInfoId);
            this.grpRegulonInfo.Items.Add(this.ddRegInfoSize);
            this.grpRegulonInfo.Items.Add(this.ddRegInfoFunction);
            this.grpRegulonInfo.Label = "regulon info column mapping";
            this.grpRegulonInfo.Name = "grpRegulonInfo";
            this.grpRegulonInfo.Visible = false;
            // 
            // ddRegInfoId
            // 
            this.ddRegInfoId.Enabled = false;
            this.ddRegInfoId.Image = global::GINtool.Properties.Resources.target;
            this.ddRegInfoId.Label = "regulon name";
            this.ddRegInfoId.Name = "ddRegInfoId";
            this.ddRegInfoId.ShowImage = true;
            this.ddRegInfoId.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddRegInfoId_SelectionChanged);
            // 
            // ddRegInfoSize
            // 
            this.ddRegInfoSize.Enabled = false;
            this.ddRegInfoSize.Image = global::GINtool.Properties.Resources.crop;
            this.ddRegInfoSize.Label = "size";
            this.ddRegInfoSize.Name = "ddRegInfoSize";
            this.ddRegInfoSize.ShowImage = true;
            this.ddRegInfoSize.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddRegInfoSize_SelectionChanged);
            // 
            // ddRegInfoFunction
            // 
            this.ddRegInfoFunction.Enabled = false;
            this.ddRegInfoFunction.Image = global::GINtool.Properties.Resources.keyboard;
            this.ddRegInfoFunction.Label = "function";
            this.ddRegInfoFunction.Name = "ddRegInfoFunction";
            this.ddRegInfoFunction.ShowImage = true;
            this.ddRegInfoFunction.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddRegInfoFunction_SelectionChanged);
            // 
            // grpColMapCategory
            // 
            this.grpColMapCategory.Items.Add(this.ddCatID);
            this.grpColMapCategory.Items.Add(this.ddCatName);
            this.grpColMapCategory.Items.Add(this.ddCatBSU);
            this.grpColMapCategory.Label = "category linkage column mappings";
            this.grpColMapCategory.Name = "grpColMapCategory";
            this.grpColMapCategory.Visible = false;
            // 
            // ddCatID
            // 
            this.ddCatID.Enabled = false;
            this.ddCatID.Image = global::GINtool.Properties.Resources.swatches;
            this.ddCatID.Label = "category id";
            this.ddCatID.Name = "ddCatID";
            this.ddCatID.ShowImage = true;
            this.ddCatID.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddCatID_SelectionChanged);
            // 
            // ddCatName
            // 
            this.ddCatName.Enabled = false;
            this.ddCatName.Image = global::GINtool.Properties.Resources.keyboard;
            this.ddCatName.Label = "description";
            this.ddCatName.Name = "ddCatName";
            this.ddCatName.ShowImage = true;
            this.ddCatName.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddCatName_SelectionChanged);
            // 
            // ddCatBSU
            // 
            this.ddCatBSU.Enabled = false;
            this.ddCatBSU.Image = global::GINtool.Properties.Resources.target;
            this.ddCatBSU.Label = "gene id";
            this.ddCatBSU.Name = "ddCatBSU";
            this.ddCatBSU.ShowImage = true;
            this.ddCatBSU.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddCatBSU_SelectionChanged);
            // 
            // grpCutOff
            // 
            this.grpCutOff.Items.Add(this.editMinPval);
            this.grpCutOff.Items.Add(this.ebLow);
            this.grpCutOff.Label = "cut-offs";
            this.grpCutOff.Name = "grpCutOff";
            this.grpCutOff.Visible = false;
            // 
            // editMinPval
            // 
            this.editMinPval.Label = "p-value";
            this.editMinPval.Name = "editMinPval";
            this.editMinPval.ScreenTip = "Define the p-value cut-off value to include in the comined report";
            this.editMinPval.Text = null;
            this.editMinPval.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditMinPval_TextChanged);
            // 
            // ebLow
            // 
            this.ebLow.Label = "fold-change";
            this.ebLow.Name = "ebLow";
            this.ebLow.ScreenTip = "Set the value for the minimum FC category";
            this.ebLow.Text = null;
            this.ebLow.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TextBox_Low_TextChanged);
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
            this.cbAscending.ScreenTip = "Set the sort direction for outputting tables and figures";
            this.cbAscending.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Ascending_Click);
            // 
            // cbDescending
            // 
            this.cbDescending.Label = "descending";
            this.cbDescending.Name = "cbDescending";
            this.cbDescending.ScreenTip = "Set the sort direction for outputting tables and figures";
            this.cbDescending.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckBox_Descending_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.tglTaskPane);
            this.group2.Items.Add(this.button2);
            this.group2.Label = "about";
            this.group2.Name = "group2";
            this.group2.Visible = false;
            // 
            // tglTaskPane
            // 
            this.tglTaskPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.tglTaskPane.Enabled = false;
            this.tglTaskPane.Image = global::GINtool.Properties.Resources.clipboard;
            this.tglTaskPane.Label = "show/hide manual";
            this.tglTaskPane.Name = "tglTaskPane";
            this.tglTaskPane.ShowImage = true;
            this.tglTaskPane.Visible = false;
            this.tglTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tglTaskPane_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::GINtool.Properties.Resources.chat;
            this.button2.Label = "About";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
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
            this.grpFilter.ResumeLayout(false);
            this.grpFilter.PerformLayout();
            this.grpFocus.ResumeLayout(false);
            this.grpFocus.PerformLayout();
            this.grpTable.ResumeLayout(false);
            this.grpTable.PerformLayout();
            this.grpPlot.ResumeLayout(false);
            this.grpPlot.PerformLayout();
            this.grpReference.ResumeLayout(false);
            this.grpReference.PerformLayout();
            this.grpGenesMapping.ResumeLayout(false);
            this.grpGenesMapping.PerformLayout();
            this.grpMap.ResumeLayout(false);
            this.grpMap.PerformLayout();
            this.grpRegulonInfo.ResumeLayout(false);
            this.grpRegulonInfo.PerformLayout();
            this.grpColMapCategory.ResumeLayout(false);
            this.grpColMapCategory.PerformLayout();
            this.grpCutOff.ResumeLayout(false);
            this.grpCutOff.PerformLayout();
            this.grpDirection.ResumeLayout(false);
            this.grpDirection.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebLow;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCutOff;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editMinPval;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetOperonFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPlot;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUseCategories;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btPlot;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectCatFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearCatFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCatFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDistribution;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbClustered;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkRegulon;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbCombined;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbSummary;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUseOperons;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUsePValues;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUseFoldChanges;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDta;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDirection;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbAscending;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbDescending;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbUseRegulons;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetRegulonLinkageFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitBtnGenesFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectGenesFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenesFileSelected;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbGenesFileMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpGenesMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddGnsName;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbRegulonMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddGenesBSU;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddGenesFunction;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddGenesDescription;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator8;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbNoFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFocus;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox dbTblRanking;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpColMapCategory;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddCatID;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddCatName;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddCatBSU;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbCategoryMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator9;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox cbRegInfoColumnMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRegInfoFileName;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpRegulonInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddRegInfoId;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddRegInfoSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddRegInfoFunction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearGenInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearRegulonInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tglTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
    }

    partial class ThisRibbonCollection
    {
        internal GinRibbon GinRibbon
        {
            get { return this.GetRibbon<GinRibbon>(); }
        }
    }
}
