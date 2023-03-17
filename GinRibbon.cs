#undef CLICK_CHART // check to include clickable chart and events.. only if object storage is an option.

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;
using static GINtool.ES_Extensions;
using stat_dict = System.Collections.Generic.Dictionary<string, double>;
using dataset_dict = System.Collections.Generic.Dictionary<string, GINtool.DataItem>;
using rank_dict = System.Collections.Generic.Dictionary<string, int>;
using dict_rank = System.Collections.Generic.Dictionary<int, string>;
using lib_dict = System.Collections.Generic.Dictionary<string, string[]>;
using Accord.Statistics.Distributions.Univariate;
using System.Text;
using System.Security.Cryptography;
//certificate CdF7RoqS9KXLvWtk6OZf chk


namespace GINtool
{
    using gsea_dict = System.Collections.Generic.Dictionary<string, GINtool.S_GSEA>;

    /// <summary>
    /// The main class of the Excel Addin
    /// </summary>

    public partial class GinRibbon
    {

        /// <value>The last used folder for an input file.</value>        
        string gLastFolder = "";

        /// <value>The flag that registers which data to update.</value>        
        byte gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;

        /// <value>The list in which the tasks are registered.</value>
        readonly List<TASKS> gTasks = new List<TASKS>();

        int gMaxGenesPerOperon = 1;
        /// <value>The main table that contains all gene info</value>
        SysData.DataTable gGenesWB = null;
        /// <value>The main table containing the regulon data.</value>
        SysData.DataTable gRegulonWB = null; // RegulonData .. rename later
        /// <value>The main table containing the operon data</value>
        SysData.DataTable gRefOperonsWB = null;
        /// <value>The main table containing the category data</value>
        SysData.DataTable gCategoriesWB = null;
        /// <value>The main table containing the regulon info data</value>
        SysData.DataTable gRegulonInfoWB = null;

        /// <value>gGeneColNames contains the column names of the genes information file</value>
        private string[] gGenesColNames = new string[] { };
        /// <value>gRegulonColNames contains the column names of the regulon file</value>
        private string[] gRegulonColNames = new string[] { };
        /// <value>gCategoryColNames contains the column names of the categories file</value>
        private string[] gCategoryColNames = new string[] { };
        /// <value>gOperonColNames contains the columns names of the operon file</value>
        private string[] gOperonColNames = new string[] { };
        /// <value>gRegulonInfoColNames contains the columns names of the regulon info file</value>
        private string[] gRegulonInfoColNames = new string[] { };

        #region ES related variables
        dataset_dict gDataSetDict = new dataset_dict();
        lib_dict gCategoryDict = new Dictionary<string, string[]>();
        lib_dict gRegulonDict = new Dictionary<string, string[]>();
        lib_dict gCombinedDict = new Dictionary<string, string[]>();
        Hashtable gFgseaHash = new Hashtable();
        Hashtable gGSEAHash = new Hashtable();
        Dictionary<string, string> gBSU_gene_dict = new Dictionary<string, string>();        
        stat_dict gES_signature = new stat_dict();
        stat_dict gES_signature_ordered = new stat_dict();
        dict_rank gES_map_signature = new dict_rank();
        rank_dict gES_signature_map = new rank_dict();
        string[] gES_signature_genes = new string[] { };
        double[] gES_sigvalues = new double[] { };
        double[] gES_abs_signature = new double[] { };
        int gES_key;
        #endregion
        //readonly string gCategoryGeneColumn = "locus_tag"; // the fixed column name that refers to the genes inthe category csv file
        Excel.Application gApplication = null;

        Excel.Workbook gActiveWorkbook = null;

        /// <value>the main list of all association types listed in the main regulon table</value>
        static List<string> gAvailItems = null;

        /// <value>the main list of items that the user defined as having a up-regulated association with a gene</value>
        static List<string> gUpItems = null;

        /// <value>the main list of items that the user defined as having a down-regulated association with a gene</value>
        static List<string> gDownItems = null;

        List<int> gExcelErrorValues = null;

        /// <value>a string that represents the previously selected range of BSU codes</value>
        string gOldRangeBSU = "";
        /// <value>a string that represents the previously selected range of P-values</value>
        string gOldRangeP = "";
        /// <value>a string that represents the previously selected range of FC</value>
        string gOldRangeFC = "";

        Excel.Range gRangeBSU;
        Excel.Range gRangeFC;
        Excel.Range gRangeP;

        List<BsuLinkedItems> gList = null;
        /// <value>Contains the usage info of the regulons</value>
        SysData.DataTable gRegulonTable = null;
        /// <value>Contains the usage info of the categories</value>
        SysData.DataTable gCategoryTable = null;
        SysData.DataTable gBestTable = null;


        bool gRegulonFileSelected = false;
        bool gCategoryFileSelected = false;
        bool gGenesFileSelected = false;

        bool gOperonFileSelected = false;
        bool gRegulonInfoFileSelected = false;


        Properties.Settings gSettings = null;

        private bool UseCategoryData()
        {
            return Properties.Settings.Default.useCat;
        }

        /// <summary>
        /// Obtain the value for the property from the default settings
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        private List<string> PropertyItems(string property)
        {
            StringCollection myCol = (StringCollection)Properties.Settings.Default[property];

            if (myCol != null)
                return myCol.Cast<string>().ToList();

            return new List<string>();
        }

        /// <summary>
        /// Store the value or values of property to the default settings
        /// </summary>
        /// <param name="property"></param>
        /// <param name="aValue"></param>
        private void StoreValue(string property, List<string> aValue)
        {

            StringCollection collection = new StringCollection();
            collection.AddRange(aValue.ToArray());

            Properties.Settings.Default[property] = collection;
        }

        /// <summary>
        /// From a table get the distinct records for the itmes listed in Columns
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Columns"></param>
        /// <returns></returns>
        private SysData.DataTable GetDistinctRecords(SysData.DataTable dt, string[] Columns)
        {
            return dt.DefaultView.ToTable(true, Columns.Distinct().ToArray());
        }

        /// <summary>
        /// Find the records in the main Regulon table where the ID (=BSU column, locus_tag) = value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private SysData.DataRow[] LookupRegulon(string value)
        {
            // needs to be replaced by genes table entry
            SysData.DataRow[] filteredRows = gRegulonWB.Select(string.Format("[{0}] LIKE '%{1}%'", Properties.Settings.Default.referenceBSU, value));

            // copy data to temporary table
            SysData.DataTable dt = gRegulonWB.Clone();
            foreach (SysData.DataRow dr in filteredRows)
                dt.ImportRow(dr);
            // return only unique values
            SysData.DataTable dt_unique = GetDistinctRecords(dt, gRegulonColNames);
            return dt_unique.Select();
        }


        /// <summary>
        /// Find the records in the main Category table where the ID (=BSU column, locus tag) = value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private SysData.DataRow[] LookupCategory(string value)
        {
            SysData.DataRow[] filteredRows = gCategoriesWB.Select(string.Format("[locus_tag] = '{0}'", value));

            // copy data to temporary table
            SysData.DataTable dt = gCategoriesWB.Clone();
            foreach (SysData.DataRow dr in filteredRows)
                dt.ImportRow(dr);
            // return only unique values
            // SysData.DataTable dt_unique = GetDistinctRecords(dt, gCategoryColNames);
            SysData.DataTable dt_unique = GetDistinctRecords(dt, new string[] { });
            return dt_unique.Select();
        }


        /// <summary>
        /// Find the records in the main Category table where the ID (=BSU column, locus tag) = value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private SysData.DataRow[] LookupGeneInfo(string value)
        {
            // needs to be replaced by genes table entry
            SysData.DataRow[] filteredRows = gGenesWB.Select(string.Format("[{0}] LIKE '%{1}%'", Properties.Settings.Default.genesBSUColumn, value));

            // copy data to temporary table
            SysData.DataTable dt = gGenesWB.Clone();
            foreach (SysData.DataRow dr in filteredRows)
                dt.ImportRow(dr);
            // return only unique values
            SysData.DataTable dt_unique = GetDistinctRecords(dt, gGenesColNames);
            return dt_unique.Select();
        }


        /// <summary>
        /// Enable/disable the buttons and labels at the start.
        /// </summary>
        /// <param name="enable"></param>

        private void InitFields(bool enable = false)
        {
            btnSelect.Enabled = enable;
            btApply.Enabled = enable;

            // genes items

            ddGenesBSU.Enabled = enable;
            ddGenesDescription.Enabled = enable;
            ddGenesFunction.Enabled = enable;
            ddGnsName.Enabled = enable;

            // regulon items

            ddBSU.Enabled = enable;
            ddGene.Enabled = enable;
            ddRegulon.Enabled = enable;
            ddDir.Enabled = enable;

            // category items

            ddCatID.Enabled = enable;
            ddCatName.Enabled = enable;
            ddCatBSU.Enabled = enable;

            // regulon info items

            ddRegInfoFunction.Enabled = enable;
            ddRegInfoId.Enabled = enable;
            ddRegInfoSize.Enabled = enable;

            btPlot.Enabled = enable;
            cbUseCategories.Enabled = enable;
            cbMapping.Enabled = enable;
            cbSummary.Enabled = enable;
            cbCombined.Enabled = enable;
            cbUseOperons.Enabled = enable;
            cbUsePValues.Enabled = enable;
            cbUseFoldChanges.Enabled = enable;
            cbNoFilter.Enabled = enable;
            toggleButton1.Enabled = true;
            cbAscending.Enabled = enable;
            cbDescending.Enabled = enable;
            cbUseRegulons.Enabled = enable;
            cbUseCategories.Enabled = enable;




        }


        /// <summary>
        /// Load the last known settings stored in the persitent default.settings
        /// </summary>
        private void LoadPersistentSettings()
        {
            btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;

            if (btnRegulonFileName.Label.Length > 0 & btnRegulonFileName.Label != "not defined yet")
            {
                try
                {
                    System.IO.FileInfo fInfo = new System.IO.FileInfo(btnRegulonFileName.Label);
                    gLastFolder = fInfo.DirectoryName;
                    if (LoadRegulonDataColumns())
                        Fill_RegulonDropDownBoxes();
                }
                catch (Exception ex)
                {
                    gApplication.StatusBar.Text = ex.Message;
                    // show error dialog here
                }
            }

            btnGenesFileSelected.Label = Properties.Settings.Default.genesFileName;
            if (btnGenesFileSelected.Label.Length > 0)
            {
                try
                {
                    System.IO.FileInfo fInfo = new System.IO.FileInfo(btnGenesFileSelected.Label);
                    gLastFolder = fInfo.DirectoryName;
                    if (LoadGenesDataColumns())
                        Fill_GenesDropDownBoxes();
                }
                catch (Exception ex)
                {
                    gApplication.StatusBar.Text = ex.Message;
                }
            }

            btnOperonFile.Label = Properties.Settings.Default.operonFile;
            if (btnOperonFile.Label.Length > 0)
            {
                try
                {
                    System.IO.FileInfo fInfo = new System.IO.FileInfo(btnOperonFile.Label);
                    gLastFolder = fInfo.DirectoryName;
                    //if (LoadOperonDataColumns())
                    //    Fill_OperonDropDownBoxes();
                }
                catch (System.Exception ex)
                {
                    gApplication.StatusBar.Text = ex.Message;

                }
            }

            btnCatFile.Label = Properties.Settings.Default.categoryFile;
            if (btnCatFile.Label.Length > 0)
            {
                try
                {
                    System.IO.FileInfo fInfo = new System.IO.FileInfo(btnCatFile.Label);
                    gLastFolder = fInfo.DirectoryName;
                    if (LoadCategoryDataColumns())
                        Fill_CategoryDropDownBoxes();
                }
                catch (Exception ex)
                {
                    gApplication.StatusBar.Text = ex.Message;

                }
            }


            //if(gSettings.operonFile.Length ==0 & gSettings.operonSheet.Length>0)

            //if (Properties.Settings.Default.categoryFile.Length == 0 & Properties.Settings.Default.referenceFile.Length > 0)
            //{
            //    cbUseCategories.Checked = false;
            //    cbUseRegulons.Checked = true;
            //    Properties.Settings.Default.useCat = false;
            //}

            btnRegInfoFileName.Label = gSettings.regulonInfoFIleName;
            if (btnRegInfoFileName.Label.Length > 0) // check this with merge from home 17/03/2022
                try
                {
                    System.IO.FileInfo fInfo = new System.IO.FileInfo(btnRegInfoFileName.Label);
                    gLastFolder = fInfo.DirectoryName;
                    if (LoadRegulonInfoDataColumns())
                        Fill_RegulonInfoDropDownBoxes();
                }
                catch (Exception ex)
                {
                    gApplication.StatusBar.Text = ex.Message;

                }

            cbDescending.Checked = !Properties.Settings.Default.sortAscending;
            cbAscending.Checked = Properties.Settings.Default.sortAscending;

            cbGSEAFC.Checked = true; // Properties.Settings.Default.gseaFC;
            cbGSEAP.Checked = false; // !Properties.Settings.Default.gseaFC;
            

            cbGenesFileMapping.Checked = false; // Properties.Settings.Default.genesMappingVisible;
            cbRegulonMapping.Checked = false; // Properties.Settings.Default.regulonMappingVisible;
            cbCategoryMapping.Checked = false;


            //operonMappingVisible

            chkRegulon.Checked = Properties.Settings.Default.regPlot;
            cbVolcano.Checked = Properties.Settings.Default.vcPlot;
            cbMapping.Checked = Properties.Settings.Default.tblMap;
            cbSummary.Checked = Properties.Settings.Default.tblSummary;
            cbCombined.Checked = Properties.Settings.Default.tblCombine;
            cbUseOperons.Checked = Properties.Settings.Default.useOperons;

            cbClustered.Checked = Properties.Settings.Default.catPlot;
            cbDistribution.Checked = Properties.Settings.Default.distPlot;

            //cbUseCategories.Checked = Properties.Settings.Default.useCat;
            //cbUseRegulons.Checked = !Properties.Settings.Default.useCat;

            cbUsePValues.Checked = Properties.Settings.Default.use_pvalues;
            cbUseFoldChanges.Checked = Properties.Settings.Default.use_foldchange;
            cbNoFilter.Checked = (cbUsePValues.Checked == false) && (cbUseFoldChanges.Checked == false);

            // load the up/down definitions 

            gAvailItems = PropertyItems("directionMapUnassigned");
            gUpItems = PropertyItems("directionMapUp");
            gDownItems = PropertyItems("directionMapDown");

            cbUseRegulons.Checked = gSettings.useRegulons & (!gSettings.useOperons & !gSettings.useCat) & (gDownItems.Count > 0 | gUpItems.Count > 0);
            cbUseCategories.Checked = gSettings.useCat & (!gSettings.useOperons & !gSettings.useRegulons);
            cbUseOperons.Checked = gSettings.useOperons & (!gSettings.useRegulons & !gSettings.useCat);


            AdjustFocusChecks();

        }

        /// <summary>
        /// The initial load procedure of the Add-in. Initialize fields and labels depending on last known settings
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void GinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            gApplication = Globals.ThisAddIn.GetExcelApplication();
            // set the static application for plot routines
            PlotRoutines.theApp = gApplication;

            gSettings = Properties.Settings.Default;
            InitFields();

            // run this line to mimic first time installation 23-03
            // gSettings.Reset();

            LoadPersistentSettings();
            EnableOutputOptions(false);

            gExcelErrorValues = ((int[])Enum.GetValues(typeof(ExcelUtils.CVErrEnum))).ToList();

            gCategoryFileSelected = System.IO.File.Exists(Properties.Settings.Default.categoryFile);
            gRegulonFileSelected = System.IO.File.Exists(Properties.Settings.Default.referenceFile);
            gGenesFileSelected = System.IO.File.Exists(Properties.Settings.Default.genesFileName);
            gOperonFileSelected = System.IO.File.Exists(Properties.Settings.Default.operonFile);
            gRegulonInfoFileSelected = System.IO.File.Exists(Properties.Settings.Default.regulonInfoFIleName);

            btLoad.Enabled = System.IO.File.Exists(gSettings.referenceFile) | System.IO.File.Exists(gSettings.categoryFile) | System.IO.File.Exists(gSettings.genesFileName) | System.IO.File.Exists(gSettings.regulonInfoFIleName);

        }


        private void AdjustFocusChecks()
        {
            if (!cbUseCategories.Enabled)
            {
                cbUseCategories.Checked = false;
                gSettings.useCat = false;
            }

            if (!cbUseOperons.Enabled)
            {
                cbUseOperons.Checked = false;
                gSettings.useOperons = false;
            }
            if (!cbUseRegulons.Enabled)
            {
                cbUseRegulons.Checked = false;
                gSettings.useRegulons = false;
            }
        }


        /// <summary>
        /// Enable/disable possible output buttons (i.e. table or charts).
        /// </summary>
        /// <param name="enable"></param>
        private void EnableOutputOptions(bool enable)
        {
            ebLow.Enabled = enable;
            editMinPval.Enabled = enable;

            cbMapping.Enabled = enable;
            cbSummary.Enabled = enable;
            cbCombined.Enabled = enable;

            cbClustered.Enabled = enable;
            cbDistribution.Enabled = enable;
            chkRegulon.Enabled = enable;
            cbVolcano.Enabled = enable;

            cbUseCategories.Enabled = enable && gCategoriesWB != null; //gCategoryFileSelected &&
            cbUseRegulons.Enabled = enable && (gRegulonWB != null && (gDownItems.Count > 0 | gUpItems.Count > 0)); //(gRegulonFileSelected 
            cbUseOperons.Enabled = enable && gRefOperonsWB != null;

            cbUsePValues.Enabled = enable;
            cbUseFoldChanges.Enabled = enable;
            cbNoFilter.Enabled = enable;

            cbAscending.Enabled = enable;
            cbDescending.Enabled = enable;

            AdjustFocusChecks();

        }

        /// <summary>
        /// Helper function to get the active cell from the current worksheet
        /// </summary>
        /// <returns></returns>

        private Excel.Range GetActiveCell()
        {
            if (gApplication != null)
            {
                try { return (Excel.Range)gApplication.Selection; }
                catch { return null; }

            }
            return null;
        }

        /// <summary>
        /// Return the active worksheet that is selected. If the active sheet is a chart then show warning message.
        /// </summary>
        /// <returns></returns>

        private Excel.Worksheet GetActiveSheet()
        {
            if (gApplication != null)
            {
                if (gApplication.ActiveSheet is Excel.Chart)
                {
                    MessageBox.Show("Please activate data sheet and select columns with data");
                    return null;
                }
                try
                {
                    return (Excel.Worksheet)gApplication.ActiveSheet;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return null;
        }

        /// <summary>
        /// Determine if the cell contains an integer, else return error
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private bool IsErrorCell(object obj)
        {

            return (obj is Int32 @int) && gExcelErrorValues.Contains(@int);
        }

        /// <summary>
        /// The main routine to map the data to a list of genes with their associated FC, p-values etc.. and a list of regulons
        /// </summary>
        /// <param name="theCells"></param>
        /// <returns>A list of data genes</returns>
        private List<BsuLinkedItems> AugmentWithRegulonData(List<BsuLinkedItems> theInputData)
        {            

            AddTask(TASKS.AUGMENTING_WITH_REGULON_DATA);

            // create a view and sort it to improve querying performance..
            DataView _regulonView = new DataView(gRegulonWB);
            _regulonView.Sort = Properties.Settings.Default.referenceBSU;

            // loop of the number of rows in rangeBSU

            foreach (BsuLinkedItems _it in theInputData)
            {                
                if ((_it.BSU.Length > 0) & !(gRegulonWB is null))
                {
                    // find the entries that are linked by the same gene
                    SysData.DataRow[] results = LookupRegulon(_it.BSU);

                    // loop over the entries (=regulons) found
                    for (int r = 0; r < results.Length; r++)
                    {

                        // check for existence if mapped to regulon
                        string item = results[r][Properties.Settings.Default.referenceRegulon].ToString();
                        string direction = results[r][Properties.Settings.Default.referenceDIR].ToString();


                        if (item.Length > 0)
                        {
                            if (gUpItems.Contains(direction))
                            {
                                _it.REGULON_UP.Add(r);
                                _it.Regulons.Add(new RegulonItem(item, "UP"));
                            }

                            if (gDownItems.Contains(direction))
                            {
                                _it.Regulons.Add(new RegulonItem(item, "DOWN"));
                                _it.REGULON_DOWN.Add(r);
                            }

                            if (!gUpItems.Contains(direction) & !gDownItems.Contains(direction))
                            {
                                _it.REGULON_UNKNOWN_DIR.Add(r);
                                _it.Regulons.Add(new RegulonItem(item, "NOT DEFINED"));
                            }

                        }                        
                    }
                }
            }

            //foreach (BsuLinkedItems _it in theInputData)
            //{               
            //    if (_it.GeneName!="")
            //        gRegulonDict.Add(_it.GeneName, _it.Regulons.Select(r => r.Name).ToArray());
            //}

            RemoveTask(TASKS.AUGMENTING_WITH_REGULON_DATA);
            return theInputData;

        }

        /// <summary>
        /// The main routine to map the data to a list of genes with their associated FC, p-values etc.. and a possibly a list of categories.. 
        /// </summary>
        /// <param name="theCells"></param>
        /// <returns>A list of data genes</returns>
        private List<BsuLinkedItems> AugmentWithCategoryData(List<BsuLinkedItems> theInputData)
        {

            AddTask(TASKS.AUGMENTING_WITH_CATEGORY_DATA);

            DataView _catView = new DataView(gCategoriesWB);
            _catView.Sort = "locus_tag"; // Properties.Settings.Default.catBSUColum;

            // loop of the number of rows in rangeBSU

            foreach (BsuLinkedItems _it in theInputData)
            {

                if ((_it.BSU.Length > 0) & !(gCategoriesWB is null))
                {
                    // find the entries that are linked by the same gene                    
                    SysData.DataRow[] results = LookupCategory(_it.BSU);
                    foreach (DataRow row in results)
                    {
                        string[] c1 = new string[] { row["cat1"].ToString(), row["cat2"].ToString(), row["cat3"].ToString(), row["cat4"].ToString(), row["cat5"].ToString() };
                        string catName = "";
                        foreach (string s in c1)
                        {
                            if (s.Length > 0)
                                catName = s;
                        }

                        string genID = row["locus_tag"].ToString();
                        string catID = row["catid_short"].ToString();

                        CategoryItem _lCat = new CategoryItem(catName, catID, genID);
                        _it.Categories.Add(_lCat);

                    }

                }
            }
            
            //foreach (BsuLinkedItems _it in theInputData)
            //{
            //    if (_it.GeneName != "")
            //        gCategoryDict.Add(_it.GeneName, _it.Categories.Select(r => r.catID).ToArray());
            //}


            RemoveTask(TASKS.AUGMENTING_WITH_CATEGORY_DATA);
            return theInputData;

        }



        /// <summary>
        /// Get the address in string format of a specified range
        /// </summary>
        /// <param name="rng"></param>
        /// <returns></returns>
        public string RangeAddress(Excel.Range rng)
        {
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }
        /// <summary>
        /// Get the addres in string format of a specified cell
        /// </summary>
        /// <param name="sht"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public string CellAddress(Excel.Worksheet sht, int row, int col)
        {
            return RangeAddress(sht.Cells[row, col]);
        }

        /// <summary>
        /// Get the string that should be displayed on the status bar, depending on the task that is running
        /// </summary>
        /// <param name="task"></param>
        /// <returns></returns>
        private string GetStatusTask(TASKS task)
        {
            return taks_strings[(int)task];
        }


        /// <summary>
        /// Set the text of the status bar
        /// </summary>
        /// <param name="activeTask"></param>
        private void SetStatus(TASKS activeTask)
        {
            gApplication.StatusBar = GetStatusTask(activeTask);
            if (activeTask != TASKS.READY)
            {
                gApplication.ScreenUpdating = false;
                gApplication.DisplayAlerts = false;
                gApplication.EnableEvents = false;
            }
            else
            {
                gApplication.ScreenUpdating = true;
                gApplication.DisplayAlerts = true;
                gApplication.EnableEvents = true;
            }

        }

        /// <summary>
        ///  Add the task to the list of tasks. The order is last in, first out (LIFO)
        /// </summary>
        /// <param name="newTask"></param>
        private void AddTask(TASKS newTask)
        {
            gTasks.Add(newTask);
            SetStatus(newTask);
        }

        /// <summary>
        /// Remove the task after completion or error and set to ready if no more tasks are performed
        /// </summary>
        /// <param name="taskReady"></param>
        private void RemoveTask(TASKS taskReady)
        {
            gTasks.Remove(taskReady);
            if (gTasks.Count == 0 || gTasks[0] == TASKS.READY)
                SetStatus(TASKS.READY);
            else
                SetStatus(gTasks.Last());
        }

        /// <summary>
        /// Detemine if a selected range is different from what previously selected
        /// </summary>
        /// <returns></returns>
        private bool InputHasChanged()
        {

            bool changed = false;
            if (gOldRangeBSU != gRangeBSU.Address.ToString())
            {
                gOldRangeBSU = gRangeBSU.Address.ToString();
                changed = true;
            }

            if (gOldRangeP != gRangeP.Address.ToString())
            {
                gOldRangeP = gRangeP.Address.ToString();
                changed = true;
            }

            if (gOldRangeFC != gRangeFC.Address.ToString())
            {
                gOldRangeFC = gRangeFC.Address.ToString();
                changed = true;
            }

            return changed;
        }




        /// <summary>
        /// Copy the text in tables to a worksheet using a single assignment to an Excel.Range
        /// </summary>
        /// <param name="dt">data table</param>
        /// <param name="sheet">existing worksheet</param>
        /// <param name="firstRow">first row in worksheet</param>
        /// <param name="firstCol">first column in worksheet</param>
        /// <param name="lastRow">last row in worksheet</param>
        /// <param name="lastCol">last column in worksheet</param>
        private void FastDtToExcel(System.Data.DataTable dt, Excel.Worksheet sheet, int firstRow, int firstCol, int lastRow, int lastCol)
        {
            Excel.Range top = sheet.Cells[firstRow, firstCol];
            Excel.Range bottom = sheet.Cells[lastRow, lastCol];
            Excel.Range all = (Excel.Range)sheet.get_Range(top, bottom);

            object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
                for (int j = 0; j < dt.Columns.Count; j++)
                    arrayDT[i, j] = dt.Rows[i][j];
            all.Value = arrayDT;

        }

        /// <summary>
        /// Copy the cells green if value is positive and red if value in table is negative
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastRow"></param>
        /// <param name="lastCol"></param>
        private void ColorCells(System.Data.DataTable dt, Excel.Worksheet sheet, int firstRow, int firstCol, int lastRow, int lastCol)
        {
            AddTask(TASKS.COLOR_CELLS);

            Excel.Range top = sheet.Cells[firstRow, firstCol];
            Excel.Range bottom = sheet.Cells[lastRow, lastCol];
            Excel.Range all = (Excel.Range)sheet.get_Range(top, bottom);


            for (int r = 0; r < dt.Rows.Count; r++)
            {
                SysData.DataRow clrRow = dt.Rows[r];
                for (int c = 0; c < clrRow.ItemArray.Length; c++)
                {
                    Excel.Range lR = all.Cells[r + 1, c + 1];
                    if (Int32.TryParse(clrRow[c].ToString(), out int val))
                    {
                        if (val == 1)
                            lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

                        if (val == -1)
                            lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                    }
                }
            }

            RemoveTask(TASKS.COLOR_CELLS);
        }

        /// <summary>
        /// Return a list of sheet that already exist
        /// </summary>
        /// <returns></returns>
        private List<string> ListSheets()
        {
            // get a list of all sheet names
            List<string> _sheets = new List<string>();

            foreach (var sheet in gApplication.Sheets)
            {
                if (sheet is Excel.Chart _c)
                    _sheets.Add(_c.Name);
                else
                    _sheets.Add(((Excel.Worksheet)sheet).Name);
            }

            return _sheets;

        }

        /// <summary>
        /// Create a new non-existing sheet name that starts with a given prefix
        /// </summary>
        /// <param name="wsBase"></param>
        /// <returns></returns>
        private int NextWorksheet(string wsBase)
        {
            // create a sheetname starting with wsBase
            List<string> currentSheets = ListSheets();

            string sheetName = wsBase.Replace("Plot", "Tab");
            string chartName = wsBase.Replace("Tab", "Plot");

            int s = 1;
            while (currentSheets.Contains(string.Format("{0}{1}", chartName, s)) || currentSheets.Contains(string.Format("{0}{1}", sheetName, s)))
                s += 1;

            return s;
        }

        /// <summary>
        /// Determine fist possible suffix for a set of worksheets.
        /// </summary>
        /// <param name="aSheet"></param>
        /// <param name="wsBase"></param>
        /// <returns></returns>
        private int FindSheetNames(string[] wsBase)
        {
            // create a sheetname starting with wsBase
            List<string> currentSheets = ListSheets();
            int s = 1;


            while (true)
            {
                List<bool> aList = new List<bool>();

                for (int i = 0; i < wsBase.Length; i++)
                    aList.Add(currentSheets.Contains(string.Format("{0}_{1}", wsBase[i], s)));

                if (!aList.Contains(true))
                    break;
                s++;
            }

            return s;
        }

        /// <summary>
        /// Rename a newly created worksheet with a given prefix.
        /// </summary>
        /// <param name="aSheet"></param>
        /// <param name="wsBase"></param>
        /// <returns></returns>
        private int RenameWorksheet(object aSheet, string wsBase)
        {
            // create a sheetname starting with wsBase
            List<string> currentSheets = ListSheets();
            int s = 1;
            while (currentSheets.Contains(string.Format("{0}_{1}", wsBase, s)))
                s += 1;

            if (aSheet is Excel.Worksheet _w)
            {
                _w.Name = string.Format("{0}_{1}", wsBase, s);
            }

            return s;
        }


        /// <summary>
        /// Return a list of genes and their FCs that are linked by a single operon
        /// </summary>
        /// <param name="opid"></param>
        /// <param name="lLst"></param>
        /// <returns></returns>
        private (List<string>, List<double>) GetOperonGenesFC(/*string operon,*/ string opid, List<BsuLinkedItems> lLst)
        {

            SysData.DataRow[] lquery = gRefOperonsWB.Select(string.Format("op_id = '{0}'", opid));

            List<string> _genes = new List<string>();
            List<double> _lfcs = new List<double>();

            foreach (DataRow row in lquery)
            {

                string lgene = row["gene"].ToString();
                _genes.Add(lgene);

                BsuLinkedItems result = lLst.Find(item => item.GeneName == lgene);
                if (result != null)
                    _lfcs.Add(result.FC);
                else
                    _lfcs.Add(Double.NaN);
            }

            return (_genes, _lfcs);
        }


        private List<BsuLinkedItems> AugmentWithGeneInfo(List<Excel.Range> theCells)
        {

            AddTask(TASKS.AUGMENTING_WITH_GENES_INFO);

            // Copy the data from the selected cells as determined in the dialog earlier.
            object[,] rangeBSU = theCells[2].Value2;
            object[,] rangeFC = theCells[1].Value2;
            object[,] rangeP = theCells[0].Value2;

            //gDataSetStat_dict.Clear();
            gDataSetDict.Clear();            
            // Initialize the list
            List<BsuLinkedItems> lList = new List<BsuLinkedItems>();

            // loop of the number of rows in rangeBSU            
            for (int _r = 1; _r <= rangeBSU.Length; _r++)
            {
                string lBSU;
                double lFC = 0;
                double lPvalue = 1;

                lBSU = rangeBSU[_r, 1].ToString();

                if (!IsErrorCell(rangeP[_r, 1]))
                    if (!Double.TryParse(rangeP[_r, 1].ToString(), out lPvalue))
                        lPvalue = 1;

                if (!IsErrorCell(rangeFC[_r, 1]))
                    if (!Double.TryParse(rangeFC[_r, 1].ToString(), out lFC))
                        lFC = 0;
                              
                // create a mapping entry .. not annotated with category data (yet!)
                BsuLinkedItems lMap = new BsuLinkedItems(lFC, lPvalue, lBSU);

                //  double check if BSU has a value 
                if ((lMap.BSU.Length > 0) & !(gGenesWB is null))
                {
                    // find the entries that are linked by the same gene                    
                    SysData.DataRow[] results = LookupGeneInfo(lMap.BSU);
                    if (results.Length > 0)
                    {
                        lMap.GeneName = results[0][Properties.Settings.Default.genesNameColumn].ToString();
                        lMap.GeneDescription = results[0][Properties.Settings.Default.genesDescriptionColumn].ToString();
                        lMap.GeneFunction = results[0][Properties.Settings.Default.genesFunctionColumn].ToString();
                    }
                }

                DataItem lItem = new DataItem
                {
                    pval = lMap.PVALUE,
                    FC = lMap.FC,
                    BSU = lMap.BSU
                };

                try
                {
                    //gDataSetStat_dict.Add(lMap.BSU, lItem.FC);                    
                    gDataSetDict.Add(lMap.BSU, lItem);          
                    // map for BSU to gene ... might be deleted later .. 
                    gBSU_gene_dict.Add(lMap.BSU, lMap.GeneName);
                }
                catch
                {
                    gApplication.StatusBar = String.Format("no information for BSU ({0}) number was found ", lMap.BSU);
                }

                lList.Add(lMap);
            }

            // focus GSEA on absolute values ONLY!! because of directionality .. 
            // gDataSetStat_dict = gDataSetDict.ToDictionary(kvp => kvp.Key, kvp => Math.Abs(kvp.Value.FC));

            RemoveTask(TASKS.AUGMENTING_WITH_GENES_INFO);

            // to be implemented (combine data with gene info table (i.e. function/description)                        
            return lList;
        }

        /// <summary>
        /// Find the item in the drop down menu by value
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private RibbonDropDownItem GetItemByValue(RibbonDropDown ctrl, string value)
        {
            RibbonDropDownItem ret = null;
            foreach (RibbonDropDownItem it in ctrl.Items)
            {
                if (it.Label == value)
                {
                    ret = it;
                    break;
                }
            }
            return ret;
        }

        /// <summary>
        /// Load the possible up/down definitions
        /// </summary>
        private void LoadDirectionOptions()
        {
            SysData.DataView view = new SysData.DataView(gRegulonWB);
            SysData.DataTable distinctValues = view.ToTable(true, Properties.Settings.Default.referenceDIR);

            foreach (SysData.DataRow row in distinctValues.Rows)
            {
                gAvailItems.Add(row.ItemArray[0].ToString());
            }
        }


        /// <summary>
        /// Fill the regulon dropdown boxes and select the last known (stored) selected value
        /// </summary>
        private void Fill_RegulonDropDownBoxes()
        {
            gApplication.EnableEvents = false;

            ddBSU.Items.Clear();
            ddRegulon.Items.Clear();
            ddGene.Items.Clear();
            ddDir.Items.Clear();

            foreach (string s in gRegulonColNames)
            {
                RibbonDropDownItem ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddBSU.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddRegulon.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddDir.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddGene.Items.Add(ddItem1);

            }

            RibbonDropDownItem ddItem = GetItemByValue(ddBSU, Properties.Settings.Default.referenceBSU);
            if (ddItem != null)
                ddBSU.SelectedItem = ddItem;
            else
                Properties.Settings.Default.referenceBSU = gRegulonColNames[0];

            ddItem = GetItemByValue(ddRegulon, Properties.Settings.Default.referenceRegulon);
            if (ddItem != null)
                ddRegulon.SelectedItem = ddItem;
            else
                Properties.Settings.Default.referenceRegulon = gRegulonColNames[0];


            ddItem = GetItemByValue(ddDir, Properties.Settings.Default.referenceDIR);
            if (ddItem != null)
                ddDir.SelectedItem = ddItem;
            else
                Properties.Settings.Default.referenceDIR = gRegulonColNames[0];


            ddItem = GetItemByValue(ddGene, Properties.Settings.Default.referenceGene);
            if (ddItem != null)
                ddGene.SelectedItem = ddItem;
            else
                Properties.Settings.Default.referenceGene = gRegulonColNames[0];

            ddBSU.Enabled = true;
            ddRegulon.Enabled = true;
            ddDir.Enabled = true;
            ddGene.Enabled = true;
            btRegDirMap.Enabled = true;

            gApplication.EnableEvents = true;


        }

        private void Fill_GenesDropDownBoxes()
        {
            gApplication.EnableEvents = false;

            ddGnsName.Items.Clear();
            ddGenesBSU.Items.Clear();
            ddGenesFunction.Items.Clear();
            ddGenesDescription.Items.Clear();


            foreach (string s in gGenesColNames)
            {
                RibbonDropDownItem ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddGnsName.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddGenesBSU.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddGenesFunction.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddGenesDescription.Items.Add(ddItem1);

            }

            RibbonDropDownItem ddItem = GetItemByValue(ddGenesBSU, gSettings.genesBSUColumn);
            if (ddItem != null)
                ddGenesBSU.SelectedItem = ddItem;
            else
                gSettings.genesBSUColumn = gGenesColNames[0];

            ddItem = GetItemByValue(ddGnsName, Properties.Settings.Default.genesNameColumn);
            if (ddItem != null)
                ddGnsName.SelectedItem = ddItem;
            else
                Properties.Settings.Default.genesNameColumn = gGenesColNames[0];

            ddItem = GetItemByValue(ddGenesFunction, Properties.Settings.Default.genesFunctionColumn);
            if (ddItem != null)
                ddGenesFunction.SelectedItem = ddItem;
            else
                Properties.Settings.Default.genesFunctionColumn = gGenesColNames[0];

            ddItem = GetItemByValue(ddGenesDescription, Properties.Settings.Default.genesDescriptionColumn);
            if (ddItem != null)
                ddGenesDescription.SelectedItem = ddItem;
            else
                Properties.Settings.Default.genesDescriptionColumn = gGenesColNames[0];

            ddGnsName.Enabled = true;
            ddGenesBSU.Enabled = true;
            ddGenesFunction.Enabled = true;
            ddGenesDescription.Enabled = true;

            gApplication.EnableEvents = true;


        }


        private void Fill_CategoryDropDownBoxes()
        {
            gApplication.EnableEvents = false;

            ddCatID.Items.Clear();
            ddCatName.Items.Clear();
            ddCatBSU.Items.Clear();

            foreach (string s in gCategoryColNames)
            {
                RibbonDropDownItem ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddCatID.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddCatName.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddCatBSU.Items.Add(ddItem1);
            }

            RibbonDropDownItem ddItem = GetItemByValue(ddCatID, Properties.Settings.Default.catCatIDColumn);
            if (ddItem != null)
                ddCatID.SelectedItem = ddItem;
            else
                gSettings.catCatIDColumn = gCategoryColNames[0];

            ddItem = GetItemByValue(ddCatName, Properties.Settings.Default.catCatDescriptionColumn);
            if (ddItem != null)
                ddCatName.SelectedItem = ddItem;
            else
                gSettings.catCatIDColumn = gCategoryColNames[0];

            ddItem = GetItemByValue(ddCatBSU, Properties.Settings.Default.catBSUColum);
            if (ddItem != null)
                ddCatBSU.SelectedItem = ddItem;
            else
                gSettings.catBSUColum = gCategoryColNames[0];

            ddCatID.Enabled = true;
            ddCatName.Enabled = true;
            ddCatBSU.Enabled = true;

            gApplication.EnableEvents = true;


        }



        private void Fill_RegulonInfoDropDownBoxes()
        {
            gApplication.EnableEvents = false;

            ddRegInfoFunction.Items.Clear();
            ddRegInfoId.Items.Clear();
            ddRegInfoSize.Items.Clear();

            foreach (string s in gRegulonInfoColNames)
            {
                RibbonDropDownItem ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddRegInfoFunction.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddRegInfoId.Items.Add(ddItem1);

                ddItem1 = Factory.CreateRibbonDropDownItem();
                ddItem1.Label = s;
                ddRegInfoSize.Items.Add(ddItem1);
            }

            RibbonDropDownItem ddItem = GetItemByValue(ddRegInfoFunction, gSettings.regInfoFunctionColumn);
            if (ddItem != null)
                ddRegInfoFunction.SelectedItem = ddItem;
            else
                gSettings.regInfoFunctionColumn = gRegulonInfoColNames[0];

            ddItem = GetItemByValue(ddRegInfoId, gSettings.regInfoIdColumn);
            if (ddItem != null)
                ddRegInfoId.SelectedItem = ddItem;
            else
                gSettings.regInfoIdColumn = gRegulonInfoColNames[0];

            ddItem = GetItemByValue(ddRegInfoSize, gSettings.regInfoSizeColumn);
            if (ddItem != null)
                ddRegInfoSize.SelectedItem = ddItem;
            else
                gSettings.regInfoSizeColumn = gRegulonInfoColNames[0];

            ddRegInfoFunction.Enabled = true;
            ddRegInfoId.Enabled = true;
            ddRegInfoSize.Enabled = true;

            gApplication.EnableEvents = true;


        }

        /// <summary>
        /// Reset main variables to initial values
        /// </summary>
        private void ResetTables()
        {          
         
            gRegulonTable = null;
            gCategoryTable = null;

            EnableOutputOptions(false);

            btApply.Enabled = false;
            btPlot.Enabled = false;

            btLoad.Enabled = gGenesWB == null;
            btnSelect.Enabled = false;

            ClearWorkbookVariables();

        }

        private void ClearWorkbookVariables()
        {
            gList = null;
            gOldRangeBSU = "";
            gOldRangeFC = "";
            gOldRangeP = "";
            ClearGSEAVariables();
            EnableOutputOptions(false);

        }


        /// <summary>
        /// Load FC values from last known settings
        /// </summary>

        private void LoadFCDefaults()
        {
            ebLow.Text = Properties.Settings.Default.fcLOW.ToString();
            editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();
        }



        /// <summary>
        /// Flag that defintion for BSU (=regulon code) has changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void DropDown_BSU_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceBSU = ddBSU.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadRegulonData();
            EnableSelectButton();
            EnableFocusItems();
        }

        /// <summary>
        /// Flag that Regulon name identifier has changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_Regulon_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceRegulon = ddRegulon.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadRegulonData();
            EnableSelectButton();
            EnableFocusItems();
        }

        private void ShowUpDownMappingDialog()
        {
            if (gAvailItems.Count == 0 && gUpItems.Count == 0 && gDownItems.Count == 0)
                LoadDirectionOptions();

            dlgUpDown dlgUD = new dlgUpDown(gAvailItems, gUpItems, gDownItems);
            dlgUD.ShowDialog();

            StoreValue("directionMapUnassigned", gAvailItems);
            StoreValue("directionMapUp", gUpItems);
            StoreValue("directionMapDown", gDownItems);

            if (gUpItems.Count > 0 | gDownItems.Count > 0)
                gApplication.StatusBar = false;

            EnableSelectButton();
            EnableFocusItems();

            SetFlags(UPDATE_FLAGS.ALL);

        }

        /// <summary>
        /// The definitions for up/down regulations have been changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_RegulonDirectionMap_Click(object sender, RibbonControlEventArgs e)
        {
            ShowUpDownMappingDialog();
           
        }

        /// <summary>
        /// The column mapping identifier has changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_RegulonDirection_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            LoadRegulonData();
            EnableSelectButton();
            Properties.Settings.Default.referenceDIR = ddDir.SelectedItem.Label;
            gAvailItems.Clear();
            gUpItems.Clear();
            gDownItems.Clear();
            try
            {
                LoadDirectionOptions();
                ShowUpDownMappingDialog();
            }
            catch (Exception ex)
            {
                gApplication.StatusBar = ex.Message;
            }
                        
            EnableFocusItems();
            SetFlags(UPDATE_FLAGS.ALL);
        }

        /// <summary>
        /// Routine to check if changes made to textbox are ok, if not reset to previous value.
        /// </summary>
        /// <param name="bx">the box that was editeds</param>

        private void ValidateTextBoxData(RibbonEditBox bx)
        {

            bool low = false;
            if (bx.Equals(ebLow))
                low = true;
            // can still add range checks e.g. high > mid > low  

            var s = bx.Text.ToString(CultureInfo.InvariantCulture);
            s = s.Replace(',', '.');

            if (double.TryParse(s,NumberStyles.Any,CultureInfo.InvariantCulture, out double val))
            {
                // set the text value to what is parsed
                bx.Text = val.ToString(CultureInfo.InvariantCulture);
                if (low)
                    Properties.Settings.Default.fcLOW = val;
                SetFlags(UPDATE_FLAGS.FC_dependent);
            }
            else
            {
                if (low)
                    ebLow.Text = Properties.Settings.Default.fcLOW.ToString();
            }
        }

        /// <summary>
        /// Check changes in textbox Low
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_Low_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ValidateTextBoxData(ebLow);
        }


        /// <summary>
        /// The main routine after the load button has been selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Load_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;

            bool _reload = true;
            if (gGenesWB != null && (gCategoriesWB != null || gRegulonInfoWB != null || gRefOperonsWB != null))
                _reload = MessageBox.Show("Really overwrite existing data?") == DialogResult.OK;


            if (_reload && (LoadGenesData() && (LoadRegulonData() | LoadCategoryData()) | LoadOperonData()))
            {
                gRegulonInfoFileSelected = LoadRegulonInfoData();
                //gOperonFileSelected = LoadOperonData();

                if (gRegulonFileSelected)
                {
                    if (gDownItems.Count == 0 && gUpItems.Count == 0 && gAvailItems.Count == 0)
                        LoadDirectionOptions();

                    if (gAvailItems.Count > 0 & (gDownItems.Count == 0 & gUpItems.Count == 0))
                        gApplication.StatusBar = "Select defintions of up and down regulation first before running regulon augmentation!";

                }

                toggleButton1.Enabled = true;
                LoadFCDefaults();
                ResetTables();

                btLoad.Enabled = false;
                btnSelect.Enabled = true;
            }
            gApplication.EnableEvents = true;
        }


        /// <summary>
        /// The main routine after Plot has been selected to create the selection of plots
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void Button_Plot_Click(object sender, RibbonControlEventArgs e)
        {
            if (!(Properties.Settings.Default.catPlot || Properties.Settings.Default.regPlot || Properties.Settings.Default.distPlot || Properties.Settings.Default.vcPlot))
            {
                MessageBox.Show("Please select at least one plot to generate");
                return;
            }

            if (gRegulonFileSelected && (gRegulonTable is null || NeedsUpdate(UPDATE_FLAGS.TRegulon)))
            {

                gRegulonTable = CreateRegulonUsageTable(GetDataSelection());
                UnSetFlags(UPDATE_FLAGS.TRegulon);
            }
            if (gCategoryFileSelected && (gCategoryTable is null || NeedsUpdate(UPDATE_FLAGS.TCategory)))
            {

                gCategoryTable = CreateCategoryUsageTable(GetDataSelection());
                UnSetFlags(UPDATE_FLAGS.TCategory);
            }



            if ((Properties.Settings.Default.catPlot || Properties.Settings.Default.regPlot || Properties.Settings.Default.vcPlot)) //& gNeedsUpdate.Check(UPDATE_FLAGS.PCat))
            {
                dlgTreeView dlg = new dlgTreeView(categoryView: cbUseCategories.Checked, spreadingOptions: Properties.Settings.Default.catPlot, rankingOptions: Properties.Settings.Default.regPlot, volcanoOptions:Properties.Settings.Default.vcPlot);

                if (gCategoriesWB != null && cbUseCategories.Checked)
                {
                    dlg.populateTree(gCategoriesWB);
                }
                if (gRegulonWB != null && !cbUseCategories.Checked)
                {
                    dlg.populateTree(gRegulonWB, cat: false);
                }

                if (dlg.ShowDialog() == DialogResult.OK)
                {

                    if (Properties.Settings.Default.regPlot && gList != null)
                        RankingPlot(dlg.GetSelection(), UseCategoryData() ? gCategoryTable : gRegulonTable);

                    if (Properties.Settings.Default.catPlot && gList != null)
                        SpreadingPlot(dlg.GetSelection(), topTenFC: dlg.getTopFC(), outputTable: dlg.selectTableOutput());

                    if (Properties.Settings.Default.vcPlot && gList != null)
                        VolcanoPlot(dlg.GetSelection(), UseCategoryData() ? gCategoryTable : gRegulonTable, maxExtreme:dlg.getExtremeP());

                }
            }

            //if (Properties.Settings.Default.distPlot)
            //{
            //    if (gList != null)
            //    {
            //        DistributionPlot(GetDataSelection());
            //    }
            //}
        }

        /// <summary>
        /// The routine after apply has been selected to created the worksheets.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void Button_Apply_Click(object sender, RibbonControlEventArgs e)
        {


            if (!(gSettings.useOperons || gSettings.useCat || gSettings.useRegulons))
            {
                MessageBox.Show("Please select at least one output table to generate");
                return;
            }

            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            if (NoUpdate())
                return;

            // combined info should contain best table info

            if (gRegulonFileSelected && (gRegulonTable is null || NeedsUpdate(UPDATE_FLAGS.TRegulon) || gRegulonTable.Rows.Count == 0))
            {

                gRegulonTable = CreateRegulonUsageTable(GetDataSelection());
                UnSetFlags(UPDATE_FLAGS.TRegulon);
            }
            if (gCategoryFileSelected && (gCategoryTable is null || NeedsUpdate(UPDATE_FLAGS.TCategory)))
            {

                gCategoryTable = CreateCategoryUsageTable(GetDataSelection());
                UnSetFlags(UPDATE_FLAGS.TCategory);
            }

            CreateBestDataTable(GetDataSelection()); //, gSettings.tblMap);

            //if (Properties.Settings.Default.tblOperon && gRefOperonsWB!=null) // can combine table/sheet because it's a quick routine
            //{
            //    SysData.DataTable tblOperon = CreateOperonTable(GetDataSelection());
            //    CreateOperonSheet(tblOperon);
            //    UnSetFlags(UPDATE_FLAGS.TOperon);
            //}

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }



        /// <summary>
        /// Set the update status for flag f
        /// </summary>
        /// <param name="f"></param>
        private void SetFlags(UPDATE_FLAGS f)
        {
            gNeedsUpdate = (byte)(gNeedsUpdate | (byte)f);
        }

        /// <summary>
        /// Remove the update status for flag f
        /// </summary>
        /// <param name="f"></param>
        private void UnSetFlags(UPDATE_FLAGS f)
        {
            gNeedsUpdate = (byte)(gNeedsUpdate & (byte)~f);
        }

        /// <summary>
        /// Check if flag f needs updating
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        private bool NeedsUpdate(UPDATE_FLAGS f)
        {
            return gNeedsUpdate.Check(f);
        }

        /// <summary>
        /// Check if any flag needs updating
        /// </summary>
        /// <returns></returns>
        private bool AnyUpdate()
        {
            return gNeedsUpdate.Any();
        }

        /// <summary>
        /// Check if no flag needs updating
        /// </summary>
        /// <returns></returns>
        private bool NoUpdate()
        {
            return gNeedsUpdate.None();
        }


        /// <summary>
        /// Create the data elements that hold the genes and FCs per category. Repeat the process those elements that have negative FCs and positive FCs.
        /// </summary>
        /// <param name="dataView"></param>
        /// <param name="cat_Elements"></param>
        /// <param name="topTenFC"></param>
        /// <param name="topTenP"></param>
        /// <returns></returns>
        private element_fc CatElements2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements, int topTenFC = -1)
        {

            //List<element_fc> element_Fcs = new List<element_fc>();
            element_fc element_Fcs = new element_fc();
            SysData.DataView dataViewCat = gCategoriesWB.AsDataView();


            List<summaryInfo> _All = new List<summaryInfo>();
            List<summaryInfo> _Pos = new List<summaryInfo>();
            List<summaryInfo> _Neg = new List<summaryInfo>();

            List<string> chk_Genes = new List<string>();

            stat_dict pvalues_all = new stat_dict();
            stat_dict pvalues_pos = new stat_dict();
            stat_dict pvalues_neg = new stat_dict();          

            foreach (cat_elements ce in cat_Elements)
            {
                string categories = string.Join(",", ce.elements.ToArray());
                categories = string.Join(",", categories.Split(',').Select(x => $"'{x}'"));
                dataViewCat.RowFilter = String.Format("catid_short in ({0})", categories);

                HashSet<string> genes = new HashSet<string>();
                foreach (DataRow _row in dataViewCat.ToTable().Rows)
                {
                    //genes.Add(_row[Properties.Settings.Default.catBSUColum].ToString());
                    genes.Add(_row["locus_tag"].ToString());
                }

                string genesFormat = string.Join(",", genes.ToArray());
                genesFormat = string.Join(",", genesFormat.Split(',').Select(x => $"'{x}'"));
                // GENE_ID moet ergens gedefinieerd worden


                dataView.RowFilter = String.Format("Gene_ID in ({0})", genesFormat);


                if (dataView.Count > 0) // set this to true to output all results.. also the zeros
                {
                    SysData.DataTable _dt = dataView.ToTable(true, "Gene_ID", "FC", "Pvalue");

                    summaryInfo __All = new summaryInfo();
                    summaryInfo __Pos = new summaryInfo();
                    summaryInfo __Neg = new summaryInfo();

                    __All.catName = ce.catName;
                    __Pos.catName = ce.catName;
                    __Neg.catName = ce.catName;

                    __All.catNameFormat = string.Format("{0} ({1})", ce.catName, _dt.Rows.Count);
                    __Pos.catNameFormat = string.Format("{0} ({1})", ce.catName, _dt.Rows.Count);
                    __Neg.catNameFormat = string.Format("{0} ({1})", ce.catName, _dt.Rows.Count);


                    List<double> _fcsA = new List<double>();
                    List<string> _genesA = new List<string>();
                    List<double> _pvaluesA = new List<double>();


                    List<double> _fcsP = new List<double>();
                    List<string> _genesP = new List<string>();
                    List<double> _pvaluesP = new List<double>();

                    List<double> _fcsN = new List<double>();
                    List<string> _genesN = new List<string>();
                    List<double> _pvaluesN = new List<double>();

                    if (_dt.Rows.Count > 0)
                    {

                        for (int i = 0; i < _dt.Rows.Count; i++)
                        {

                            double fc = (double)_dt.Rows[i]["FC"];

                            chk_Genes.Add(_dt.Rows[i]["Gene_ID"].ToString());
                            _genesA.Add(_dt.Rows[i]["Gene_ID"].ToString());
                            _fcsA.Add(fc);
                            _pvaluesA.Add(double.Parse(_dt.Rows[i]["Pvalue"].ToString()));

                            if (fc > 0) // was if (fc >= 0) , 19-10-22
                            {
                                _genesP.Add(_dt.Rows[i]["Gene_ID"].ToString());
                                _fcsP.Add(fc);
                                _pvaluesP.Add(double.Parse(_dt.Rows[i]["Pvalue"].ToString()));
                            }
                            else if (fc < 0)
                            {
                                _genesN.Add(_dt.Rows[i]["Gene_ID"].ToString());
                                _fcsN.Add(fc);
                                _pvaluesN.Add(double.Parse(_dt.Rows[i]["Pvalue"].ToString()));
                            }
                        }                     
                    }

                    __Pos.fc_average = _fcsP.Count > 0 ? _fcsP.Average() : Double.NaN;
                    __Neg.fc_average = _fcsN.Count > 0 ? _fcsN.Average() : Double.NaN;
                    __All.fc_average = _fcsA.Count > 0 ? _fcsA.Average() : Double.NaN;

                    __Pos.fc_values = _fcsP.Count > 0 ? _fcsP.ToArray() : new double[0];
                    __Neg.fc_values = _fcsN.Count > 0 ? _fcsN.ToArray() : new double[0];
                    __All.fc_values = _fcsA.Count > 0 ? _fcsA.ToArray() : new double[0];


                    __Pos.genes = _genesP.Count > 0 ? _genesP.ToArray() : new string[0];
                    __Neg.genes = _genesN.Count > 0 ? _genesN.ToArray() : new string[0];
                    __All.genes = _genesA.Count > 0 ? _genesA.ToArray() : new string[0];


                    __Pos.p_values = _pvaluesP.Count > 0 ? _pvaluesP.ToArray() : new double[0];
                    __Neg.p_values = _pvaluesN.Count > 0 ? _pvaluesN.ToArray() : new double[0];
                    __All.p_values = _pvaluesA.Count > 0 ? _pvaluesA.ToArray() : new double[0];

                    __All.fc_mad = _fcsA.Count > 0 ? _fcsA.mad() : Double.NaN;
                    __Pos.fc_mad = _fcsP.Count > 0 ? _fcsP.mad() : Double.NaN;
                    __Neg.fc_mad = _fcsN.Count > 0 ? _fcsN.mad() : Double.NaN;

                    (__Pos.es, __Pos.p_average) = CalcES(_genesP);
                    (__Neg.es, __Neg.p_average) = CalcES(_genesN); 
                    (__All.es, __All.p_average) = CalcES(_genesA);

                    pvalues_all.Add(ce.catName, __All.p_average);
                    pvalues_pos.Add(ce.catName, __Pos.p_average);
                    pvalues_neg.Add(ce.catName, __Neg.p_average);

                    __Pos.p_mad = _pvaluesP.Count > 0 ? _pvaluesP.mad() : Double.NaN;
                    __Neg.p_mad = _pvaluesN.Count > 0 ? _pvaluesN.mad() : Double.NaN;
                    __All.p_mad = _pvaluesA.Count > 0 ? _pvaluesA.mad() : Double.NaN;

                    _All.Add(__All);
                    _Pos.Add(__Pos);
                    _Neg.Add(__Neg);
                }
            }

            #region FDR

            stat_dict pvalues_fdr_all = new stat_dict();
            stat_dict pvalues_fdr_pos = new stat_dict();
            stat_dict pvalues_fdr_neg = new stat_dict();


            double[] _pvalues_fdr_all = fdr_correction(pvalues_all.Values.ToArray());
            double[] _pvalues_fdr_neg = fdr_correction(pvalues_neg.Values.ToArray());
            double[] _pvalues_fdr_pos = fdr_correction(pvalues_pos.Values.ToArray());

            int cnt = 0;
            foreach (string key in pvalues_all.Keys)
                pvalues_fdr_all.Add(key, _pvalues_fdr_all[cnt++]);

            cnt = 0;
            foreach (string key in pvalues_neg.Keys)
                pvalues_fdr_neg.Add(key, _pvalues_fdr_neg[cnt++]);

            cnt = 0;
            foreach (string key in pvalues_pos.Keys)
                pvalues_fdr_pos.Add(key, _pvalues_fdr_pos[cnt++]);

            _All = _All.Select(d => new summaryInfo()
            {
                catName = d.catName,
                es = d.es,
                catNameFormat = d.catNameFormat,
                p_values = d.p_values,
                fc_values = d.fc_values,
                p_fdr = pvalues_fdr_all[d.catName],
                p_average = d.p_average,
                fc_average = d.fc_average,
                p_mad = d.p_mad,
                fc_mad = d.fc_mad,
                genes = d.genes,
                best_gene_percentage = d.best_gene_percentage
            }).ToList();

            _Neg = _Neg.Select(d => new summaryInfo()
            {
                catName = d.catName,
                es = d.es,
                catNameFormat = d.catNameFormat,
                p_values = d.p_values,
                fc_values = d.fc_values,
                p_fdr = pvalues_fdr_neg[d.catName],
                p_average = d.p_average,
                fc_average = d.fc_average,
                p_mad = d.p_mad,
                fc_mad = d.fc_mad,
                genes = d.genes,
                best_gene_percentage = d.best_gene_percentage
            }).ToList();

            _Pos = _Pos.Select(d => new summaryInfo()
            {
                catName = d.catName,
                es = d.es,
                catNameFormat = d.catNameFormat,
                p_values = d.p_values,
                fc_values = d.fc_values,
                p_fdr = pvalues_fdr_pos[d.catName],
                p_average = d.p_average,
                fc_average = d.fc_average,
                p_mad = d.p_mad,
                fc_mad = d.fc_mad,
                genes = d.genes,
                best_gene_percentage = d.best_gene_percentage
            }).ToList();


            #endregion FDR

            element_Fcs.All = _All;
            element_Fcs.Activated = _Pos;
            element_Fcs.Repressed = _Neg;

            if (element_Fcs.All is null)
                return element_Fcs;
            


            if (topTenFC > 0) // if top FC is selected only select the top X values using absolute FC values. The default order is descending
            {
                // only useful if top N selected is smaller then total number of items
                if (topTenFC < element_Fcs.All.Count)
                {
                    double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                    var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                    List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                    element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).Where(x => x.fc_values.Length > 0).ToList();
                    element_Fcs.All = element_Fcs.All.GetRange(0, topTenFC);
                }

                // if selected, reverse the order
                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
        
            // finally, reverse order.. 
            double[] ___values = element_Fcs.All.Select(x => x.fc_average).ToArray();
            var sortedElementsFC = (!Properties.Settings.Default.sortAscending) ? ___values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : 
                ___values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();

            element_Fcs.All = sortedElementsFC.Select(x => element_Fcs.All[x.Value]).ToList();                

            return element_Fcs;
        }

        /// <summary>
        /// Create the data elements that hold the regulons and associated genes. This is also split in situations where regulon is inhibited or activated. 
        /// </summary>
        /// <param name="dataView"></param>
        /// <param name="cat_Elements"></param>
        /// <param name="topTenFC"></param>
        /// <param name="topTenP"></param>        
        /// <returns></returns>

        private element_fc Regulons2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements, int topTenFC = -1)
        {
            element_fc element_Fcs = new element_fc();

            List<summaryInfo> _All = new List<summaryInfo>();
            List<summaryInfo> _Act = new List<summaryInfo>();
            List<summaryInfo> _Rep = new List<summaryInfo>();

            stat_dict pvalues_all = new stat_dict();
            stat_dict pvalues_pos = new stat_dict();
            stat_dict pvalues_neg = new stat_dict();
           

            stat_dict pvalues_fdr_all = new stat_dict();
            stat_dict pvalues_fdr_pos = new stat_dict();
            stat_dict pvalues_fdr_neg = new stat_dict();

           

            foreach (cat_elements el in cat_Elements)
            {
                dataView.RowFilter = String.Format("Regulon='{0}'", el.catName);

                SysData.DataTable _dataTable = dataView.ToTable();

                if (_dataTable.Rows.Count > 0) // set this to true to select all rows, also the empty ones
                {

                    // find genes for the regulon/category

                    summaryInfo __All = new summaryInfo();
                    summaryInfo __Act = new summaryInfo();
                    summaryInfo __Rep = new summaryInfo();


                    __All.catName = el.catName;
                    __Act.catName = el.catName;
                    __Rep.catName = el.catName;

                    __All.catNameFormat = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                    __Act.catNameFormat = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                    __Rep.catNameFormat = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);

                    //__All.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                    //__Act.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                    //__Rep.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);

                    //total
                    List<double> _fcsT = new List<double>();
                    List<string> _genesT = new List<string>();
                    List<double> _pvaluesT = new List<double>();

                    //activatedd
                    List<double> _fcsA = new List<double>();
                    List<string> _genesA = new List<string>();
                    List<double> _pvaluesA = new List<double>();

                    // repressed
                    List<double> _fcsR = new List<double>();
                    List<string> _genesR = new List<string>();
                    List<double> _pvaluesR = new List<double>();


                    if (_dataTable.Rows.Count > 0)
                    {

                        for (int i = 0; i < _dataTable.Rows.Count; i++)
                        {
                            double fc = (double)_dataTable.Rows[i]["FC"];
                            string _geneName = _dataTable.Rows[i]["Gene_ID"].ToString();
                            _genesT.Add(_geneName);
                            _fcsT.Add(fc);
                            _pvaluesT.Add(double.Parse(_dataTable.Rows[i]["Pvalue"].ToString()));
                        }


                        // DataRow[] _inhibited = _dataTable.Select("(FC<0 AND DIR>0) OR (FC>0 AND DIR<=0) ");
                        DataRow[] _inhibited = _dataTable.Select("(FC<0 AND DIR>0) OR (FC>0 AND DIR<0) ");
                        for (int i = 0; i < _inhibited.Length; i++)
                        {
                            double fc = (double)_inhibited[i]["FC"];
                            string _geneName = _inhibited[i]["Gene_ID"].ToString();
                            _genesR.Add(_geneName);
                            _fcsR.Add(fc);
                            _pvaluesR.Add(double.Parse(_inhibited[i]["Pvalue"].ToString()));
                        }


                        //DataRow[] _activated = _dataTable.Select("(FC>0 AND DIR>0) OR (FC<0 AND DIR<=0) ");
                        DataRow[] _activated = _dataTable.Select("(FC>0 AND DIR>0) OR (FC<0 AND DIR<0) ");
                        for (int i = 0; i < _activated.Length; i++)
                        {
                            double fc = (double)_activated[i]["FC"];
                            string _geneName = _activated[i]["Gene_ID"].ToString();
                            _genesA.Add(_geneName);
                            _fcsA.Add(fc);
                            _pvaluesA.Add(double.Parse(_activated[i]["Pvalue"].ToString()));
                        }                       

                    }
              
                    __Act.fc_average = _fcsA.Count > 0 ? _fcsA.AbsAverage() : Double.NaN;
                    __Rep.fc_average = _fcsR.Count > 0 ? -_fcsR.AbsAverage() : Double.NaN;
                    __All.fc_average = _fcsT.Count > 0 ? _fcsT.Average() : Double.NaN;

                    __Act.fc_values = _fcsA.Count > 0 ? _fcsA.ToArray() : new double[0];// { 0 };
                    __Rep.fc_values = _fcsR.Count > 0 ? _fcsR.ToArray() : new double[0];// { 0 };
                    __All.fc_values = _fcsT.Count > 0 ? _fcsT.ToArray() : new double[0];// { 0 };

                    __Act.fc_mad = _fcsA.Count > 0 ? _fcsA.AbsMad() : Double.NaN;
                    __Rep.fc_mad = _fcsR.Count > 0 ? _fcsR.AbsMad() : Double.NaN;
                    __All.fc_mad = _fcsT.Count > 0 ? _fcsT.mad() : Double.NaN;

                    __Act.genes = _genesA.Count > 0 ? _genesA.ToArray() : new string[0]; // { "" };
                    __Rep.genes = _genesR.Count > 0 ? _genesR.ToArray() : new string[0];// { "" };
                    __All.genes = _genesT.Count > 0 ? _genesT.ToArray() : new string[0];// { "" };


                    __Act.p_values = _pvaluesA.Count > 0 ? _pvaluesA.ToArray() : new double[0];// { };
                    __Rep.p_values = _pvaluesR.Count > 0 ? _pvaluesR.ToArray() : new double[0];// { };
                    __All.p_values = _pvaluesT.Count > 0 ? _pvaluesT.ToArray() : new double[0];// { };
                   
                    (__All.es, __All.p_average) = CalcES(_genesT);
                    (__Act.es, __Act.p_average) = CalcES(_genesA);
                    (__Rep.es, __Rep.p_average) = CalcES(_genesR);


                    pvalues_all.Add(el.catName, __All.p_average);
                    pvalues_pos.Add(el.catName, __Act.p_average);
                    pvalues_neg.Add(el.catName, __Rep.p_average);


                    __Act.p_mad = _pvaluesA.Count > 0 ? _pvaluesA.mad() : Double.NaN;
                    __Rep.p_mad = _pvaluesR.Count > 0 ? _pvaluesR.mad() : Double.NaN;
                    __All.p_mad = _pvaluesT.Count > 0 ? _pvaluesT.mad() : Double.NaN;

                    _All.Add(__All);
                    _Act.Add(__Act);
                    _Rep.Add(__Rep);

                }
            }

            #region FDR
            double[] _pvalues_fdr_all = fdr_correction(pvalues_all.Values.ToArray());
            double[] _pvalues_fdr_neg = fdr_correction(pvalues_neg.Values.ToArray());
            double[] _pvalues_fdr_pos = fdr_correction(pvalues_pos.Values.ToArray());

            int cnt = 0;
            foreach (string key in pvalues_all.Keys)
                pvalues_fdr_all.Add(key, _pvalues_fdr_all[cnt++]);

            cnt = 0;
            foreach (string key in pvalues_neg.Keys)
                pvalues_fdr_neg.Add(key, _pvalues_fdr_neg[cnt++]);

            cnt = 0;
            foreach (string key in pvalues_pos.Keys)
                pvalues_fdr_pos.Add(key, _pvalues_fdr_pos[cnt++]);
           
            _All = _All.Select(d => new summaryInfo()
            {
                catName = d.catName,
                es = d.es,
                catNameFormat = d.catNameFormat,
                p_values = d.p_values,
                fc_values = d.fc_values,
                p_fdr = pvalues_fdr_all[d.catName],
                p_average = d.p_average,
                fc_average = d.fc_average,
                p_mad = d.p_mad,
                fc_mad = d.fc_mad,
                genes = d.genes,
                best_gene_percentage = d.best_gene_percentage
            }).ToList();

            _Rep = _Rep.Select(d => new summaryInfo()
            {
                catName = d.catName,
                es = d.es,
                catNameFormat = d.catNameFormat,
                p_values = d.p_values,
                fc_values = d.fc_values,
                p_fdr = pvalues_fdr_neg[d.catName],
                p_average = d.p_average,
                fc_average = d.fc_average,
                p_mad = d.p_mad,
                fc_mad = d.fc_mad,
                genes = d.genes,
                best_gene_percentage = d.best_gene_percentage
            }).ToList();

            _Act = _Act.Select(d => new summaryInfo()
            {
                catName = d.catName,
                es = d.es,
                catNameFormat = d.catNameFormat,
                p_values = d.p_values,
                fc_values = d.fc_values,
                p_fdr = pvalues_fdr_pos[d.catName],
                p_average = d.p_average,
                fc_average = d.fc_average,
                p_mad = d.p_mad,
                fc_mad = d.fc_mad,
                genes = d.genes,
                best_gene_percentage = d.best_gene_percentage
            }).ToList();

            #endregion

            element_Fcs.All = _All;
            element_Fcs.Activated = _Act;
            element_Fcs.Repressed = _Rep;



            if (element_Fcs.All is null)
                return element_Fcs;


            if (topTenFC > 0) // top X FC is based on abs average FC.
            {
                // only useful if top N selected is smaller then total number of items
                if (topTenFC < element_Fcs.All.Count)
                {
                    double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                    var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                    //List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                    // remove elements with no genes associated
                    element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList().Where(x => x.fc_values.Length > 0).ToList();
                    element_Fcs.All = element_Fcs.All.GetRange(0, topTenFC);
                }
            }           
            //
            
            double[] ___values = element_Fcs.All.Select(x => x.fc_average).ToArray();
            var _sortedElements = (!Properties.Settings.Default.sortAscending) ? ___values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : ___values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();
            //List<int> _sortedIndex = _sortedElements.Select(x => x.Value).ToList();
            element_Fcs.All = _sortedElements.Select(x => element_Fcs.All[x.Value]).ToList();

            //if (!Properties.Settings.Default.sortAscending)
                //    element_Fcs.All.Reverse();
            

            return element_Fcs;
        }


        /// <summary>
        /// Get the sorted FCs with the corresponding index in the data table
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        private (List<double>, List<int>) SortedFoldChanges(SysData.DataTable dataTable)
        {
            List<double> _values = new List<double>();
            foreach (SysData.DataRow row in dataTable.Rows)
            {
                _values.Add(row.Field<double>("FC"));
            }

            double[] __values = _values.ToArray();
            var sortedGenes = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();


            List<double> sortedGenesValues = sortedGenes.Select(x => x.Key).ToList();
            List<int> sortedGenesInt = sortedGenes.Select(x => x.Value).ToList();
            return (sortedGenesValues, sortedGenesInt);
        }



        /// <summary>
        /// In the subtiwiki data there can be similar categories and/or regulons. This assures a unique selection of those.
        /// </summary>
        /// <param name="elements"></param>
        /// <returns></returns>
        private List<cat_elements> GetUniqueElements(List<cat_elements> elements)
        {
            List<cat_elements> result = new List<cat_elements>();
            foreach (cat_elements el in elements)
            {
                cat_elements elo = result.GetCatElement(el.catName);
                if (elo.catName == null)
                    result.Add(el);
                else
                {
                    System.Console.WriteLine("Hallo");
                }

            }

            return result;
        }

        /// <summary>
        /// Utility to Strip the word regulon from a string.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string StripText(string str)
        {
            string name = str;
            string[] names = name.Split('(');
            int hit = names[0].ToUpper().IndexOf("REGULON");
            string newname = hit == -1 ? names[0] : names[0].Substring(0, hit);
            return newname;
        }

        /// <summary>
        /// Sort the elements in a list of summaryInfo according to a specific field and order
        /// </summary>
        /// <param name="alist"></param>
        /// <param name="mode"></param>
        /// <param name="descending"></param>
        /// <returns></returns>
        private List<summaryInfo> SortedElements(List<summaryInfo> alist, SORTMODE mode = SORTMODE.FC, bool descending = true)
        {
            List<summaryInfo> _work = new List<summaryInfo>(alist);


            if (mode == SORTMODE.CATNAME)
            {
                string[] __values = _work.Select(x => x.catName).ToArray();
                var sortedElements = descending ? __values.Select((x, i) => new KeyValuePair<string, int>(x, i)).OrderByDescending(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<string, int>(x, i)).OrderBy(x => x.Key).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();

                alist = sortedElements.Select(x => _work[x.Value]).ToList();

                return alist;
            }
            else
            {
                double[] __values = null;
                switch (mode)
                {
                    case SORTMODE.FC:
                        __values = _work.Select(x => x.fc_average).ToArray();
                        break;
                    case SORTMODE.P:
                        //__values = _work.Select(x => x.p_average).ToArray();
                        __values = _work.Select(x => x.p_fdr).ToArray();
                        break;
                    default:
                        __values = _work.Select(x => x.fc_average).ToArray();
                        break;

                }
                var sortedElements = descending ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();

                alist = sortedElements.Select(x => _work[x.Value]).ToList();

                return alist;
            }
        }

        /// <summary>
        /// Export the category/regulon spreading data to a table that contains all genes associated with that category/regulon. 
        /// </summary>
        /// <param name="sInfo"></param>
        /// <returns></returns>
        private DataTable ElementsToExtendedTable(List<summaryInfo> sInfo)
        {
            SysData.DataTable lTable = new SysData.DataTable("ExtElements");

            string catRegLabel = Properties.Settings.Default.useCat ? "Category" : "Regulon";

            SysData.DataColumn regColumn = new SysData.DataColumn(catRegLabel, Type.GetType("System.String"));
            SysData.DataColumn geneColumn = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            SysData.DataColumn pvColumn = new SysData.DataColumn("Pvalue", Type.GetType("System.Double"));

            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(geneColumn);
            lTable.Columns.Add(fcColumn);
            lTable.Columns.Add(pvColumn);

            for (int r = 0; r < sInfo.Count; r++)
            {

                for (int g = 0; g < sInfo[r].genes.Count(); g++)
                {
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow[catRegLabel] = sInfo[r].catName;
                    lRow["Gene"] = sInfo[r].genes[g];
                    if (sInfo[r].fc_values.Length > 0)
                    {
                        lRow["FC"] = sInfo[r].fc_values[g];
                        lRow["Pvalue"] = sInfo[r].p_values[g];
                    }
                }
            }

            return lTable;
        }



        /// <summary>
        /// Transform the ranking info to a summarized table 
        /// </summary>
        /// <param name="elements"></param>
        /// <param name="bestMode"></param>
        /// <returns></returns>
        private DataTable ElementsToTable(List<summaryInfo> elements, bool bestMode = false)
        {

            SysData.DataTable lTable = new SysData.DataTable("Elements");

            SysData.DataColumn regColumn = new SysData.DataColumn("Name", Type.GetType("System.String"));
            SysData.DataColumn dirColumn = new SysData.DataColumn("Direction", Type.GetType("System.String"));
            SysData.DataColumn percColumn = new SysData.DataColumn("Percentage", Type.GetType("System.Double"));
            SysData.DataColumn cntColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
            SysData.DataColumn madColumn = new SysData.DataColumn("Mad", Type.GetType("System.Double"));
            SysData.DataColumn avgPColumn = new SysData.DataColumn("P_Average", Type.GetType("System.Double"));
            //SysData.DataColumn fdrPColumn = new SysData.DataColumn("P_FDR", Type.GetType("System.Double"));

            lTable.Columns.Add(regColumn);
            if (bestMode)
            {
                lTable.Columns.Add(dirColumn);
                lTable.Columns.Add(cntColumn);
                lTable.Columns.Add(percColumn);
            }
            else
                lTable.Columns.Add(cntColumn);
            lTable.Columns.Add(avgColumn);
            lTable.Columns.Add(madColumn);

            lTable.Columns.Add(avgPColumn);

            //if(bestMode)
            //    lTable.Columns.Add(fdrPColumn);

            if (bestMode && gRegulonInfoFileSelected && !gSettings.useCat)
            {
                SysData.DataColumn regInfoColumn = new SysData.DataColumn("Function", Type.GetType("System.String"));
                lTable.Columns.Add(regInfoColumn);
            }

            for (int r = 0; r < elements.Count; r++)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                string name = elements[r].catName;
                string[] names = name.Split('(');
                int hit = names[0].ToUpper().IndexOf("REGULON");
                string newname = hit == -1 ? names[0] : names[0].Substring(0, hit);

                if (bestMode && (gRegulonInfoFileSelected && gRegulonInfoWB != null) && !gSettings.useCat)
                {
                    DataView dataView = gRegulonInfoWB.DefaultView;
                    dataView.RowFilter = String.Format("[{0}] = '{1}'", gSettings.regInfoIdColumn, name);
                    if (dataView.Count > 0)
                        lRow["Function"] = dataView[0][gSettings.regInfoFunctionColumn];

                }

                lRow["Name"] = newname;
                if (bestMode)
                {
                    lRow["Direction"] = (elements[r].best_gene_percentage == 0.0) ? "not defined" : elements[r].fc_average > 0 ? "activation" : "repression";
                    lRow["Percentage"] = elements[r].best_gene_percentage;
                    //if (!(elements[r].p_fdr is Double.NaN))
                    //    lRow["P_FDR"] = elements[r].p_fdr;
                }

                lRow["Count"] = elements[r].p_values == null ? 0 : elements[r].p_values.Count();

                if (!(elements[r].fc_average is Double.NaN))
                    lRow["Average"] = elements[r].fc_average;
                if (!(elements[r].fc_mad is Double.NaN))
                    lRow["Mad"] = elements[r].fc_mad.ToString();
                if (!(elements[r].p_fdr is Double.NaN))
                    lRow["P_Average"] = elements[r].p_fdr; // instead of p_average

            }

            return lTable;

        }

        /// <summary>
        /// Select a csv file for regulon input
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSelectRegulonFile(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = gLastFolder;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.referenceFile = openFileDialog.FileName;
                    btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;
                    SpecifyRegulonWorksheets();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.referenceFile);
                    gLastFolder = fInfo.DirectoryName;

                    SelectFocusItem(FOCUS_ITEMS.REGULONS);

                    if (LoadRegulonDataColumns())
                    {
                        ResetTables();
                        Fill_RegulonDropDownBoxes();
                        if (!LoadRegulonData())
                            ShowMappingPanel(MAPPING_PANEL.REGULON_LINKAGE, true);
                    }
                }
            }
            EnableSelectButton();
            EnableFocusItems();
        }


        private void SelectFocusItem(FOCUS_ITEMS item)
        {

            gSettings.useOperons = item == FOCUS_ITEMS.OPERONS;
            cbUseOperons.Checked = item == FOCUS_ITEMS.OPERONS;

            gSettings.useCat = item == FOCUS_ITEMS.CATEGORIES;
            cbUseCategories.Checked = item == FOCUS_ITEMS.CATEGORIES;

            gSettings.useRegulons = (item == FOCUS_ITEMS.REGULONS && (gUpItems.Count > 0 || gDownItems.Count > 0));
            cbUseRegulons.Checked = (item == FOCUS_ITEMS.REGULONS && (gUpItems.Count > 0 || gDownItems.Count > 0));


        }


        /// <summary>
        /// Select csv file for input as operon info
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_SelectOperonFile_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = gLastFolder;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ResetTables();
                    Properties.Settings.Default.operonFile = openFileDialog.FileName;
                    btnOperonFile.Label = Properties.Settings.Default.operonFile;
                    SpecifyOperonSheet();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.operonFile);
                    gLastFolder = fInfo.DirectoryName;

                    SelectFocusItem(FOCUS_ITEMS.OPERONS);

                    if (LoadOperonDataColumns())
                    {
                        LoadOperonData();
                        //ResetTables();
                    }
                }
            }
            EnableSelectButton();
            EnableFocusItems();
        }

        /// <summary>
        /// The gene column mapping is changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_Gene_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceGene = ddGene.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadRegulonData();
            EnableSelectButton();
            EnableFocusItems();

        }

        /// <summary>
        /// The text in the editbox for minimum p-value has changed. Update p-value dependent calculations.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void EditMinPval_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var s = editMinPval.Text.ToString(CultureInfo.InvariantCulture);
            s = s.Replace(',', '.');

            if (double.TryParse(s,NumberStyles.Any,CultureInfo.InvariantCulture, out double val))
            {
                // set the text value to what is parsed
                editMinPval.Text = val.ToString();
                Properties.Settings.Default.pvalue_cutoff = val;

                SetFlags(UPDATE_FLAGS.P_dependent);
            }
            else
                editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();
        }

        /// <summary>
        /// Show/hide the GIN tool manual
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToggleTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            //var taskpane = TaskPaneManager.GetTaskPane("A", "GIN tool manual", () => new GINtaskpane(), SetTaskPaneVisbile);
            //taskpane.Visible = !taskpane.Visible;
        }

        /// <summary>
        /// Store the visibility status of the task pane
        /// </summary>
        /// <param name="visible"></param>
        public void SetTaskPaneVisbile(bool visible)
        {
            tglTaskPane.Checked = visible;
        }


        private bool CanAugmentWithCategoryData()
        {
            //? gCategoryFileSelected
            return gCategoriesWB != null; //&& cbUseCategories.Enabled;
        }

        private bool CanAugmentWithRegulonData()
        {
            // ? gRegulonFileSelected
            return gRegulonWB != null && (gUpItems.Count > 0 || gDownItems.Count > 0);
        }


        private bool CanAugmentWithOperonData()
        {
            // ? gRegulonFileSelected
            return gRefOperonsWB != null;
        }


        /// <summary>
        /// Reset the the operon file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_ResetOperonFile_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.operonFile = "";
            Properties.Settings.Default.operonSheet = "";
            btnOperonFile.Label = "No file selected";

            gOperonFileSelected = false;
            cbUseOperons.Checked = false;
            cbUseOperons.Enabled = false;

            gRefOperonsWB = null;

            Properties.Settings.Default.useOperons = false;
            EnableSelectButton();
        }

        /// <summary>
        /// Register the selection of ordering by FC (instead of P-value).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_OrderFC_Click(object sender, RibbonControlEventArgs e)
        {
            //Properties.Settings.Default.useSort = cbOrderFC.Checked;
            //gOrderAscending = cbOrderFC.Checked;
        }

        /// <summary>
        /// Select csv file for input as category info
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_SelectCatFile_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = gLastFolder;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.categoryFile = openFileDialog.FileName;
                    btnCatFile.Label = Properties.Settings.Default.categoryFile;
                    SpecifyCategorySheet();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.categoryFile);
                    gLastFolder = fInfo.DirectoryName;


                    SelectFocusItem(FOCUS_ITEMS.CATEGORIES);
                    //Properties.Settings.Default.useCat = true;
                    //cbUseCategories.Checked = true;

                    //cbUseRegulons.Checked = false;
                    //cbUseCategories.Enabled = false;

                    if (LoadCategoryDataColumns())
                    {
                        ResetTables();
                        Fill_CategoryDropDownBoxes();
                        if (!LoadCategoryData())
                            ShowMappingPanel(MAPPING_PANEL.CATEGORY_LINKAGE, true);
                    }

                }
            }
            EnableSelectButton();
            EnableFocusItems();
        }


        private void EnableSelectButton()
        {
            btnSelect.Enabled = gGenesWB != null && (gCategoriesWB != null | gRegulonWB != null | gRefOperonsWB != null);
            if (gList != null)
            {
                if (gCategoriesWB != null)
                    cbUseCategories.Enabled = true;
                if (gRegulonWB != null && (gUpItems.Count > 0 || gDownItems.Count > 0))
                    cbUseRegulons.Enabled = true;
            }

        }

        /// <summary>
        /// Register the choice of categories (instead of regulons).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_UseCategories_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.useCat = cbUseCategories.Checked;

            if (cbUseCategories.Checked)
            {
                cbUseOperons.Checked = false;
                cbUseRegulons.Checked = false;
                gSettings.useOperons = false;
                gSettings.useRegulons = false;
            }
            EnableFunctionButtons();

            //cbUseRegulons.Checked = !cbUseCategories.Checked;
            //cbUseOperons.Checked = !cbUseCategories.Checked;
        }

        /// <summary>
        /// Register the choice of a spreading plot.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Spreading_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.catPlot = cbClustered.Checked;
        }

        /// <summary>
        /// Register choice for distribution plot.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Distribution_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.distPlot = cbDistribution.Checked;
        }

        /// <summary>
        /// Register choice for ranking plot
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Regulon_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.regPlot = chkRegulon.Checked;
        }

        /// <summary>
        /// Show or hide the settings panel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Toggle_Settings_Click(object sender, RibbonControlEventArgs e)
        {
            ShowSettingPannels(toggleButton1.Checked);
            grpPlot.Visible = !toggleButton1.Checked;
            grpTable.Visible = !toggleButton1.Checked;
            grpDta.Visible = !toggleButton1.Checked;
            grpFocus.Visible = !toggleButton1.Checked;
            grpFilter.Visible = !toggleButton1.Checked;
           
        }

        /// <summary>
        /// Show or hide the panels defined as settings
        /// </summary>
        /// <param name="show"></param>
        private void ShowSettingPannels(bool show)
        {
            grpReference.Visible = show;

            button2.Label = "";
            button2.ShowLabel = false;
            grpGenesMapping.Visible = false;
            grpMap.Visible = false;
            grpColMapCategory.Visible = false;
            grpRegulonInfo.Visible = false;


            cbCategoryMapping.Checked = false;
            cbGenesFileMapping.Checked = false;
            cbRegulonMapping.Checked = false;

            grpCutOff.Visible = show;
            grpDirection.Visible = show;

            gbFGSEA.Visible = false; // show;

            group2.Visible = show;

        }

        /// <summary>
        /// Update the mapping panels based on selections by the user
        /// </summary>
        private void UpdateMappingPanels()
        {
            grpGenesMapping.Visible = cbGenesFileMapping.Checked;
            grpMap.Visible = cbRegulonMapping.Checked;
            grpColMapCategory.Visible = cbCategoryMapping.Checked;
            grpRegulonInfo.Visible = cbRegInfoColumnMapping.Checked;

            bool _bShowOther = !(grpGenesMapping.Visible | grpMap.Visible | grpColMapCategory.Visible | grpRegulonInfo.Visible);

            grpCutOff.Visible = _bShowOther;
            grpDirection.Visible = _bShowOther;
            gbFGSEA.Visible = false; // _bShowOther;

        }

        /// <summary>
        /// enumeration of tasks that are defined
        /// </summary>
        public enum TASKS : int
        {
            /// <value>Nothing to do</value>
            READY = 0,
            /// <value>Genes data is being read from file</value>
            LOAD_GENES_DATA,
            /// <value>Regulon data is being read from file</value>
            LOAD_REGULON_DATA,
            /// <value>Operon data is being read from file</value>
            LOAD_OPERON_DATA,
            /// <value>Category is being read from file</value>
            LOAD_CATEGORY_DATA,
            /// <value>Data is augmented with gene info</value>
            AUGMENTING_WITH_GENES_INFO,
            /// <value>Genes are being mapped to regulons</value>
            AUGMENTING_WITH_REGULON_DATA,
            /// <value>Genes are being mapped to categories</value>
            AUGMENTING_WITH_CATEGORY_DATA,
            /// <value>Read the selected data from the worksheet</value>
            READ_SHEET_DATA,
            /// <value>Import the the category data</value>
            READ_SHEET_CAT_DATA,
            /// <value>Update the main mapping table</value>
            UPDATE_MAPPED_TABLE,
            /// <value>Update the summary table</value>
            UPDATE_SUMMARY_TABLE,
            /// <value>Update the combined (summary/mapping) table</value>
            UDPATE_COMBINED_TABLE,
            /// <value>Update the operon data</value>
            UPDATE_OPERON_TABLE,  // a table for now.. should become table & graph           
            /// <value>Color the worksheet cells</value>
            COLOR_CELLS,
            /// <value>Create category plot</value>
            CATEGORY_CHART,
            /// <value>Create regulon plot</value>
            REGULON_CHART,
            /// <value>Enrichment score calibration</value>
            ES_CALIBRATION,
            /// <value>Calculate enrichment scores</value>
            ES_CALCULATION,
            /// <value>Generate volcano plot</value>
            VOLCANO_PLOT,
        };

        public string[] taks_strings = new string[] { "Ready", "Load genes data","Load regulon data", "Load operon data", "Load category data",
        "Augmenting with gene data", "Augmenting with with regulon data", "Augmenting with category data", "Read sheet data", "Read sheet categorized data",
            "Update mapping table", "Update summary table", "Update combined table", "Update operon table", "Color cells", "Create category chart",
            "Create regulon chart", "Calibrate enrichment scores", "Calculate enrichment scores","Create volcano plot"};

        private enum FOCUS_ITEMS : int
        {
            OPERONS = 0,
            REGULONS = 1,
            CATEGORIES = 2
        };

        public enum MAPPING_PANEL : int
        {
            GENE_INFO = 0,
            REGULON_LINKAGE = 1,
            REGULON_INFO = 2,
            OPERON_INFO = 3,
            CATEGORY_LINKAGE = 4
        };

        /// <summary>
        /// enumeration of binary flags that can be set/unset.
        /// </summary>
        public enum UPDATE_FLAGS : byte
        {
            TRegulon = 0b_0000_0001,
            TCategory = 0b_0000_0010,
            TOperon = 0b_0000_0100,
            TMapped = 0b_0000_1000, // kan andere bestemming krijgen
            PRegulon = 0b_0001_0000, // kan andere bestemming krijgen
            PDist = 0b_0010_0000,
            PCat = 0b_0100_0000,
            POperon = 0b_1000_0000,

            ///<value>FC dependency of multiple tables</value>
            FC_dependent = TCategory | POperon | TRegulon,
            ///<value>P-value dependency of multiple tables</value>
            P_dependent = TCategory | POperon | TRegulon,

            ///<value>everything needs to be updated</value>
            ALL = 0b_1111_1111,
            ///<value>no updating necessary</value>
            NONE = 0b_0000_0000
        };

        /// <summary>
        /// enumeration modes for sorting the data.
        /// </summary>

        private enum SORTMODE : int
        {
            FC = 0,
            P = 1,
            MADP = 2,
            NGENES = 3,
            CATNAME = 4,
        };

        /// <summary>
        /// Register choice for outputting mapping table.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Mapping_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblMap = cbMapping.Checked;
        }

        /// <summary>
        /// Register choice for outputting summary table.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Summary_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblSummary = cbSummary.Checked;
        }

        /// <summary>
        /// Register choice for outputting combined table.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void CheckBox_Combined_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblCombine = cbCombined.Checked;
        }

        /// <summary>
        /// Register choice for outputting operon table.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Operon_Click(object sender, RibbonControlEventArgs e)
        {
            gSettings.useOperons = cbUseOperons.Checked;

            if (cbUseOperons.Checked)
            {
                cbUseCategories.Checked = false;
                cbUseRegulons.Checked = false;
                gSettings.useCat = false;
                gSettings.useRegulons = false;
            }
            EnableFunctionButtons();
        }

        /// <summary>
        /// Clears the settings when category file is unset.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_ClearCatFile_Click(object sender, RibbonControlEventArgs e)
        {
            //gCatOutput = false;
            cbUseCategories.Checked = false;
            cbUseCategories.Enabled = false;
            Properties.Settings.Default.useCat = false;

            gCategoriesWB = null;
            gSettings.catBSUColum = "";
            gSettings.catCatDescriptionColumn = "";
            gSettings.catCatIDColumn = "";

            Properties.Settings.Default.categoryFile = "";
            btnCatFile.Label = "No file selected";



            ShowMappingPanel(MAPPING_PANEL.CATEGORY_LINKAGE, false);
            Fill_CategoryDropDownBoxes();


            EnableSelectButton();

        }

        /// <summary>
        /// Get the filtered dataset, based on Fold Change, P-value, combined or no filtering
        /// </summary>
        /// <returns></returns>
        private List<BsuLinkedItems> GetDataSelection()
        {

            if (cbNoFilter.Checked)
                return gList;

            List<BsuLinkedItems> bsuLinkedItems = new List<BsuLinkedItems>();

            foreach (BsuLinkedItems linkedItems in gList)
            {
                double lowFCVal = Properties.Settings.Default.fcLOW;
                double lowPVal = Properties.Settings.Default.pvalue_cutoff;
                bool acceptPvalue = Properties.Settings.Default.use_pvalues ? linkedItems.PVALUE <= lowPVal : true;
                bool acceptFC = Properties.Settings.Default.use_foldchange ? Math.Abs(linkedItems.FC) >= lowFCVal : true;
                if (acceptFC && acceptPvalue)
                    bsuLinkedItems.Add(linkedItems);

            }
            return bsuLinkedItems;
        }

        /// <summary>
        /// Register choice for p-values (instead of FCs).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void CheckBox_UsePValues_Click(object sender, RibbonControlEventArgs e)
        {
            gSettings.use_pvalues = cbUsePValues.Checked;
            if (cbUsePValues.Checked)
                cbNoFilter.Checked = false;
            else if (!cbUseFoldChanges.Checked)
                cbNoFilter.Checked = true;
            SetFlags(UPDATE_FLAGS.P_dependent);

        }

        /// <summary>
        /// Register choice to use FCs (instead of P-Values)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_UseFoldChanges_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.use_foldchange = cbUseFoldChanges.Checked;
            if (cbUseFoldChanges.Checked)
                cbNoFilter.Checked = false;
            else if (!cbUsePValues.Checked)
                cbNoFilter.Checked = true;

            SetFlags(UPDATE_FLAGS.FC_dependent);

        }

        /// <summary>
        /// No filtering was selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbNoFilter_Click(object sender, RibbonControlEventArgs e)
        {
            cbUseFoldChanges.Checked = !cbNoFilter.Checked;
            cbUsePValues.Checked = !cbNoFilter.Checked;
            Properties.Settings.Default.use_foldchange = !cbNoFilter.Checked;
            Properties.Settings.Default.use_pvalues = !cbNoFilter.Checked;

            SetFlags(UPDATE_FLAGS.FC_dependent & UPDATE_FLAGS.P_dependent);

        }

        private void EnableFocusItems()
        {
            cbUseCategories.Enabled = gGenesWB != null && CanAugmentWithCategoryData();
            cbUseOperons.Enabled = gGenesWB != null && CanAugmentWithOperonData();
            cbUseRegulons.Enabled = gGenesWB != null && CanAugmentWithRegulonData();

            //cbUseCategories.Checked = cbUseCategories.Enabled && CanAugmentWithCategoryData() ;
            //cbUseOperons.Checked = cbUseOperons.Enabled  && CanAugmentWithOperonData();
            //cbUseRegulons.Checked = cbUseRegulons.Enabled && CanAugmentWithCategoryData();

            gSettings.useCat = cbUseCategories.Checked;
            gSettings.useOperons = cbUseOperons.Checked;
            gSettings.useRegulons = cbUseRegulons.Checked;

            AdjustFocusChecks();


        }
       
        /// <summary>
        /// What to do when data is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Select_Click(object sender, RibbonControlEventArgs e)
        {

            gActiveWorkbook = gApplication.ActiveWorkbook;
            gActiveWorkbook.BeforeClose += GActiveWorkbook_BeforeClose;            
            Excel.Range theInputCells = GetActiveCell();

            dlgSelectData sd = new dlgSelectData(theInputCells)
            {
                theApp = gApplication
            };

            if (sd.ShowDialog() == DialogResult.OK)
            {
                // ranges are already corrected for header yes/no
                gRangeBSU = sd.getBSU();
                gRangeFC = sd.getFC();
                gRangeP = sd.getP();


                gList = ReadDataAndAugment();

                if (gList is null)
                    return;

                EnableOutputOptions(true);
                EnableFunctionButtons();

                UnSetFlags(UPDATE_FLAGS.TMapped);

            }

        }

        private void ClearGSEAVariables()
        {
            gBSU_gene_dict.Clear();
            gDataSetDict.Clear();
            gGSEAHash.Clear();
            gES_signature.Clear();
            gES_signature_ordered.Clear();
            gES_map_signature.Clear();
            gES_signature_map.Clear();
            gES_signature_genes = new string[] { };
            gES_sigvalues = new double[] { };
            gES_abs_signature = new double[] { };
            gES_key = 0;

            if (!(gList is null))
                gList.Clear();
        
        }

        private void GActiveWorkbook_BeforeClose(ref bool Cancel)
        {

            EnableOutputOptions(false);
            EnableFunctionButtons();

            ClearGSEAVariables();

            if (!(gActiveWorkbook is null))
            {
                gActiveWorkbook.Close();
                gActiveWorkbook = null;
            }
        }

        private void EnableFunctionButtons()
        {
            bool choiceTable = cbUseOperons.Checked || cbUseCategories.Checked || cbUseRegulons.Checked;
            bool choicePlot = cbUseCategories.Checked || cbUseRegulons.Checked;
            bool b = choiceTable && gGenesWB != null && (CanAugmentWithRegulonData() || CanAugmentWithCategoryData() || CanAugmentWithOperonData());
            btApply.Enabled = b;
            b = choicePlot && gGenesWB != null && (CanAugmentWithRegulonData() || CanAugmentWithCategoryData());
            btPlot.Enabled = b;

            // the plot option buttons
            cbClustered.Enabled = b;
            chkRegulon.Enabled = b;
            cbVolcano.Enabled = b;

        }


        /// <summary>
        /// Register preference for sorting in descending mode (instead of ascending)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Descending_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.sortAscending = !cbDescending.Checked;
            cbAscending.Checked = !cbDescending.Checked;
        }

        /// <summary>
        /// Register preference for sorting in ascending mode (instead of descending)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Ascending_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.sortAscending = cbAscending.Checked;
            cbDescending.Checked = !cbAscending.Checked;
        }

        /// <summary>
        /// Use regulon data for plotting (instead of categories)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void CheckBox_UseRegulons_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.useRegulons = cbUseRegulons.Checked;
            if (cbUseRegulons.Checked)
            {
                cbUseCategories.Checked = false;
                cbUseOperons.Checked = false;
                gSettings.useOperons = false;
                gSettings.useCat = false;
            }
            EnableFunctionButtons();
        }

        private void btnResetRegulonFile_Click(object sender, RibbonControlEventArgs e)
        {

            Properties.Settings.Default.referenceFile = "";
            btnRegulonFileName.Label = "No file selected";

            gRegulonFileSelected = false;

            gSettings.referenceGene = "";
            gSettings.referenceBSU = "";
            gSettings.referenceDIR = "";
            gSettings.referenceRegulon = "";
            gUpItems.Clear();
            gDownItems.Clear();
            gAvailItems.Clear();

            cbUseRegulons.Checked = false;
            cbUseRegulons.Enabled = false;
            gRegulonWB = null;

            ShowMappingPanel(MAPPING_PANEL.REGULON_LINKAGE, false);
            Fill_RegulonDropDownBoxes();


            EnableSelectButton();

        }


        private void ShowMappingPanel(MAPPING_PANEL aPanel, bool bShow)
        {
            // set all other checkboxes to false;
            RibbonCheckBox[] checkboxes = new RibbonCheckBox[] { cbGenesFileMapping, cbRegulonMapping, cbRegInfoColumnMapping, null, cbCategoryMapping };
            for (int i = 0; i < checkboxes.Length; i++)
            {
                if (checkboxes[i] != null && i != (int)aPanel)
                {
                    checkboxes[i].Checked = false;
                }
            }
            checkboxes[(int)aPanel].Checked = bShow;

            UpdateMappingPanels();

        }


        private void btnGenesFileMapping_Click(object sender, RibbonControlEventArgs e)
        {
            ShowMappingPanel(MAPPING_PANEL.GENE_INFO, cbGenesFileMapping.Checked);
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {

            ShowMappingPanel(MAPPING_PANEL.REGULON_LINKAGE, cbRegulonMapping.Checked);

        }


        private void cbCategoryMapping_Click(object sender, RibbonControlEventArgs e)
        {
            ShowMappingPanel(MAPPING_PANEL.CATEGORY_LINKAGE, cbCategoryMapping.Checked);
        }

        private void btnSelectGenesFile_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = gLastFolder;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.genesFileName = openFileDialog.FileName;
                    btnGenesFileSelected.Label = Properties.Settings.Default.genesFileName;
                    SpecifyGenesWorksheets();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.genesFileName);
                    gLastFolder = fInfo.DirectoryName;

                    if (LoadGenesDataColumns())
                    {
                        ResetTables();
                        Fill_GenesDropDownBoxes();
                        cbGenesFileMapping.Checked = true;

                        if (!LoadGenesData())
                            ShowMappingPanel(MAPPING_PANEL.GENE_INFO, true);

                    }
                }
            }

            EnableSelectButton();
        }

        private void ddGnsName_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.genesNameColumn = ddGnsName.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadGenesData();
            EnableSelectButton();
        }

        private void ddGenesBSU_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.genesBSUColumn = ddGenesBSU.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadGenesData();
            EnableSelectButton();
        }

        private void ddGenesFunction_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.genesFunctionColumn = ddGenesFunction.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadGenesData();
            EnableSelectButton();
        }

        private void ddGenesDescription_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.genesDescriptionColumn = ddGenesDescription.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadGenesData();
            EnableSelectButton();
        }

        private void ddCatName_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.catCatDescriptionColumn = ddCatName.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadCategoryData();
            EnableSelectButton();
        }

        private void ddCatID_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.catCatIDColumn = ddCatID.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadCategoryData();
            EnableSelectButton();
        }

        private void ddCatBSU_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.catBSUColum = ddCatBSU.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            LoadCategoryData();
            EnableSelectButton();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = gLastFolder;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    gSettings.regulonInfoFIleName = openFileDialog.FileName;
                    btnRegInfoFileName.Label = gSettings.regulonInfoFIleName;

                    SpecifyRegulonInfoSheet();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.regulonInfoFIleName);
                    gLastFolder = fInfo.DirectoryName;

                    if (LoadRegulonInfoDataColumns())
                    {
                        ResetTables();
                        Fill_RegulonInfoDropDownBoxes();
                        if (!LoadRegulonInfoData())
                            ShowMappingPanel(MAPPING_PANEL.REGULON_INFO, true);
                    }

                }
            }
        }

        private void cbRegInfoColumnMapping_Click(object sender, RibbonControlEventArgs e)
        {
            ShowMappingPanel(MAPPING_PANEL.REGULON_INFO, cbRegInfoColumnMapping.Checked);
        }

        private void ddRegInfoId_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            gSettings.regInfoIdColumn = ddRegInfoId.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            gRegulonInfoFileSelected = LoadRegulonInfoData();
            EnableSelectButton();
        }

        private void ddRegInfoSize_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            gSettings.regInfoSizeColumn = ddRegInfoSize.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            gRegulonInfoFileSelected = LoadRegulonInfoData();
            EnableSelectButton();
        }

        private void ddRegInfoFunction_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            gSettings.regInfoFunctionColumn = ddRegInfoFunction.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
            gRegulonInfoFileSelected = LoadRegulonInfoData();
            EnableSelectButton();
        }

        private void splitBtnGenesFile_Click(object sender, RibbonControlEventArgs e)
        {
            cbGenesFileMapping.Checked = !cbGenesFileMapping.Checked;
            ShowMappingPanel(MAPPING_PANEL.GENE_INFO, cbGenesFileMapping.Checked);
        }

        private void splitButton1_Click(object sender, RibbonControlEventArgs e)
        {
            cbRegulonMapping.Checked = !cbRegulonMapping.Checked;
            ShowMappingPanel(MAPPING_PANEL.REGULON_LINKAGE, cbRegulonMapping.Checked);

        }

        private void splitButton4_Click(object sender, RibbonControlEventArgs e)
        {
            cbCategoryMapping.Checked = !cbCategoryMapping.Checked;
            ShowMappingPanel(MAPPING_PANEL.CATEGORY_LINKAGE, cbCategoryMapping.Checked);

        }

        private void splitButton3_Click(object sender, RibbonControlEventArgs e)
        {
            cbRegInfoColumnMapping.Checked = !cbRegInfoColumnMapping.Checked;
            ShowMappingPanel(MAPPING_PANEL.REGULON_INFO, cbRegInfoColumnMapping.Checked);

        }

        private void btnClearGenInfo_Click(object sender, RibbonControlEventArgs e)
        {
            gGenesWB = null;
            gSettings.genesFileName = "";
            gSettings.genesSheetName = "";

            gSettings.genesBSUColumn = "";
            gSettings.genesDescriptionColumn = "";
            gSettings.genesFunctionColumn = "";
            gSettings.genesNameColumn = "";

            btnGenesFileSelected.Label = "No file selected";

            gList = null;
            gOldRangeBSU = "";
            gOldRangeFC = "";
            gOldRangeP = "";

            ShowMappingPanel(MAPPING_PANEL.GENE_INFO, false);
            Fill_GenesDropDownBoxes();

            EnableSelectButton();
        }

        private void btnClearRegulonInfo_Click(object sender, RibbonControlEventArgs e)
        {
            gSettings.regInfoIdColumn = "";
            gSettings.regInfoSizeColumn = "";
            gSettings.regInfoFunctionColumn = "";
            gSettings.regulonInfoFIleName = "";
            gSettings.regulonInfoSheet = "";
            gRegulonInfoWB = null;
            btnRegInfoFileName.Label = "No file selected";

            ShowMappingPanel(MAPPING_PANEL.REGULON_INFO, false);
            if (gRegulonInfoColNames.Length > 0)
                Fill_RegulonInfoDropDownBoxes();

        }

        private void tglTaskPane_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox1 ab = new AboutBox1();
            ab.ShowDialog();
        }

        private void cbVolcano_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.vcPlot = cbVolcano.Checked;
        }

        private void cbGSEAFC_Click(object sender, RibbonControlEventArgs e)
        {
            cbGSEAP.Checked = !cbGSEAFC.Checked;
            Properties.Settings.Default.gseaFC = cbGSEAFC.Checked;

            ClearWorkbookVariables();
        }

        private void cbGSEAP_Click(object sender, RibbonControlEventArgs e)
        {
            cbGSEAFC.Checked = !cbGSEAP.Checked;
            Properties.Settings.Default.gseaFC = cbGSEAFC.Checked;
            ClearWorkbookVariables();

        }
    }

    /// <summary>
    /// The basic struct for augmenting the input data (=FC, P-VALUE and BSU) with DIR and GENE
    /// </summary>

    public struct FC_BSU
    {
        public FC_BSU(double a, string b, int dir, double pval, string gene)
        {
            FC = a;
            BSU = b;
            DIR = dir;
            PVALUE = pval;
            GENE = gene;
        }
        public double FC { get; }
        public string BSU { get; }
        public double DIR { get; }
        public double PVALUE { get; }
        public string GENE { get; }
    }
}
