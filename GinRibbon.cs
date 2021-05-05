#undef CLICK_CHART // check to include clickable chart and events.. only if object storage is an option.

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;


namespace GINtool
{

    /// <summary>
    /// The main class of the Excel Addin
    /// </summary>

    public partial class GinRibbon
    {

        /// <value>The last used folder for an input file.</value>        
        string gLastFolder = "";
        bool gOperonOutput = false;
        bool gCatOutput = false;

        /// <value>The flag that registers which data to update.</value>        
        byte gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;

        /// <value>The list in which the tasks are registered.</value>
        readonly List<TASKS> gTasks = new List<TASKS>();

        int maxGenesPerOperon = 1;

        /// <value>The main table containing the regulon data.</value>
        SysData.DataTable gRefWB = null; // RegulonData .. rename later
        /// <value>The main table containing simple statistics per gene</value>
        SysData.DataTable gRefStats = null;
        /// <value>The main table containing the operon data</value>
        SysData.DataTable gRefOperons = null;
        /// <value>The main table containing the category data</value>
        SysData.DataTable gCategories = null;
        string[] gColNames = null;
        readonly string gCategoryGeneColumn = "locus_tag"; // the fixed column name that refers to the genes inthe category csv file
        Excel.Application gApplication = null;

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

        List<FC_BSU> gOutput = null;
        SysData.DataTable gSummary = null;
        List<BsuRegulons> gList = null;
        SysData.DataTable gCombineInfo = null;



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
            return dt.DefaultView.ToTable(true, Columns);
        }

        /// <summary>
        /// Find the records in the main Regulon table where the ID (=BSU column, locus_tag) = value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private SysData.DataRow[] Lookup(string value)
        {
            SysData.DataRow[] filteredRows = gRefWB.Select(string.Format("[{0}] LIKE '%{1}%'", Properties.Settings.Default.referenceBSU, value));

            // copy data to temporary table
            SysData.DataTable dt = gRefWB.Clone();
            foreach (SysData.DataRow dr in filteredRows)
                dt.ImportRow(dr);
            // return only unique values
            SysData.DataTable dt_unique = GetDistinctRecords(dt, gColNames);
            return dt_unique.Select();
        }

        /// <summary>
        /// Import the category data from the csv file downloaded from http://subtiwiki.uni-goettingen.de/. 
        /// This routine is specifically made for that specific data format.
        /// </summary>
        /// <returns></returns>
        private bool LoadCategoryData()
        {
            if (Properties.Settings.Default.categoryFile.Length == 0 || Properties.Settings.Default.catSheet.Length == 0)
                return false;

            AddTask(TASKS.LOAD_CATEGORY_DATA);

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.catSheet, Properties.Settings.Default.categoryFile, 1, 1);
            gCategories = new SysData.DataTable("Categories")
            {
                CaseSensitive = false
            };

            // long list of columns... make cleaner later..

            gCategories.Columns.Add("catid", Type.GetType("System.String"));
            gCategories.Columns.Add("catid_short", Type.GetType("System.String"));
            gCategories.Columns.Add("gene", Type.GetType("System.String"));
            gCategories.Columns.Add("locus_tag", Type.GetType("System.String"));
            gCategories.Columns.Add("cat1", Type.GetType("System.String"));
            gCategories.Columns.Add("cat2", Type.GetType("System.String"));
            gCategories.Columns.Add("cat3", Type.GetType("System.String"));
            gCategories.Columns.Add("cat4", Type.GetType("System.String"));
            gCategories.Columns.Add("cat5", Type.GetType("System.String"));
            gCategories.Columns.Add("cat1_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("cat2_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("cat3_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("cat4_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("cat5_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("ucat1_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("ucat2_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("ucat3_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("ucat4_int", Type.GetType("System.Int32"));
            gCategories.Columns.Add("ucat5_int", Type.GetType("System.Int32"));


            string[] lcols = new string[] { "cat1_int", "cat2_int", "cat3_int", "cat4_int", "cat5_int" };
            string[] ulcols = new string[] { "ucat1_int", "ucat2_int", "ucat3_int", "ucat4_int", "ucat5_int" };

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {
                object[] lItems = lRow.ItemArray;
                SysData.DataRow lNewRow = gCategories.Rows.Add();
                for (int i = 0; i < lItems.Length; i++)
                {
                    lNewRow["catid"] = lItems[0];
                    string[] splits = lItems[0].ToString().Split(' ');
                    lNewRow["catid_short"] = splits[splits.Count() - 1];
                    lNewRow["Gene"] = lItems[1];
                    lNewRow["locus_tag"] = lItems[2];
                    lNewRow["cat1"] = lItems[3];
                    lNewRow["cat2"] = lItems[4];
                    lNewRow["cat3"] = lItems[5];
                    lNewRow["cat4"] = lItems[6];
                    lNewRow["cat5"] = lItems[7];

                    string[] llItems = lItems[0].ToString().Split(' ')[1].Split('.');



                    for (int j = 0; j < llItems.Length; j++)
                    {
                        lNewRow[lcols[j]] = Int32.Parse(llItems[j]);
                    }

                    int offset = 0;
                    for (int j = 0; j < llItems.Length; j++)
                    {
                        lNewRow[ulcols[j]] = offset + ((Int32)lNewRow[lcols[j]]) * Math.Pow(10, 5 - j);
                        offset = (Int32)lNewRow[ulcols[j]];
                    }

                }
            }

            RemoveTask(TASKS.LOAD_CATEGORY_DATA);

            return gCategories.Rows.Count > 0;
        }

        /// <summary>
        /// Load the operon data from the specified csv file as downloaded from http://subtiwiki.uni-goettingen.de/.        
        /// </summary>
        /// <returns></returns>

        private bool LoadOperonData()
        {

            if (Properties.Settings.Default.operonFile.Length == 0 || Properties.Settings.Default.operonSheet.Length == 0)
                return false;

            AddTask(TASKS.LOAD_OPERON_DATA);

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.operonSheet, Properties.Settings.Default.operonFile, 1, 1);
            gRefOperons = new SysData.DataTable("OPERONS")
            {
                CaseSensitive = false
            };

            gRefOperons.Columns.Add("operon", Type.GetType("System.String"));
            gRefOperons.Columns.Add("gene", Type.GetType("System.String"));
            gRefOperons.Columns.Add("op_id", Type.GetType("System.Int32"));

            int _op_id = 0;

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {
                string[] lItems = lRow.ItemArray[0].ToString().Split('-');

                if (maxGenesPerOperon < lItems.Length)
                    maxGenesPerOperon = lItems.Length;

                for (int i = 0; i < lItems.Length; i++)
                {
                    SysData.DataRow lNewRow = gRefOperons.Rows.Add();
                    lNewRow["operon"] = lItems[0];
                    lNewRow["gene"] = lItems[i];
                    lNewRow["op_id"] = _op_id;
                }

                _op_id++;
            }

            RemoveTask(TASKS.LOAD_OPERON_DATA);
            return gRefOperons.Rows.Count > 0;
        }

        /// <summary>
        /// Load the main Regulon data as downloaded from http://subtiwiki.uni-goettingen.de/. 
        /// The whole add-in is written for data in that specific format!
        /// </summary>
        /// <returns></returns>
        private bool LoadData()
        {

            AddTask(TASKS.LOAD_REGULON_DATA);

            gRefWB = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.referenceSheetName, Properties.Settings.Default.referenceFile, 1, 1);
            if (gRefWB != null)
            {
                gColNames = new string[gRefWB.Columns.Count];
                int i = 0;
                foreach (SysData.DataColumn col in gRefWB.Columns)
                {
                    gColNames[i++] = col.ColumnName;
                }
                // generate database frequency table
                CreateTableStatistics();
            }

            RemoveTask(TASKS.LOAD_REGULON_DATA);
            return gRefWB != null;
        }

        /// <summary>
        /// Create a basic count / average usage table per regulon
        /// </summary>
        private void CreateTableStatistics()
        {
            List<string> lString = new List<string> { Properties.Settings.Default.referenceRegulon };
            SysData.DataTable lUnique = GetDistinctRecords(gRefWB, lString.ToArray());

            // initialize the global datatable

            gRefStats = new SysData.DataTable("tblstat");

            int totNrRows = gRefWB.Rows.Count;

            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn countColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
            gRefStats.Columns.Add(regColumn);
            gRefStats.Columns.Add(countColumn);
            gRefStats.Columns.Add(avgColumn);

            foreach (SysData.DataRow lRow in lUnique.Rows)
            {
                string lVal = lRow[Properties.Settings.Default.referenceRegulon].ToString();
                int cnt = gRefWB.Select(string.Format("{0}='{1}'", Properties.Settings.Default.referenceRegulon, lVal)).Length;
                SysData.DataRow nRow = gRefStats.Rows.Add();
                nRow["Regulon"] = lVal;
                nRow["Count"] = cnt;
                nRow["Average"] = ((double)cnt) / totNrRows;
            }
        }


        /// <summary>
        /// Enable/disable the buttons and labels at the start.
        /// </summary>
        /// <param name="enable"></param>

        private void InitFields(bool enable = false)
        {
            btnSelect.Enabled = enable;
            btApply.Enabled = enable;
            ddBSU.Enabled = enable;
            ddGene.Enabled = enable;
            ddRegulon.Enabled = enable;
            ddDir.Enabled = enable;
            btPlot.Enabled = enable;
            cbUseCategories.Enabled = enable;
            cbMapping.Enabled = enable;
            cbSummary.Enabled = enable;
            cbCombined.Enabled = enable;
            cbOperon.Enabled = enable;
            cbOrderFC.Enabled = enable;
            cbUsePValues.Enabled = enable;
            cbUseFoldChanges.Enabled = enable;
            toggleButton1.Enabled = true;
            cbAscending.Enabled = enable;
            cbDescending.Enabled = enable;
            cbUseRegulons.Enabled = enable;
        }


        /// <summary>
        /// Load the last known settings stored in the persitent default.settings
        /// </summary>
        private void LoadButtonStatus()
        {
            gApplication = Globals.ThisAddIn.GetExcelApplication();
            btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;

            if (btnRegulonFileName.Label.Length > 0)
            {
                System.IO.FileInfo fInfo = new System.IO.FileInfo(btnRegulonFileName.Label);
                gLastFolder = fInfo.DirectoryName;

            }

            btnOperonFile.Label = Properties.Settings.Default.operonFile;
            btnCatFile.Label = Properties.Settings.Default.categoryFile;
            cbOrderFC.Checked = Properties.Settings.Default.useSort;
            cbDescending.Checked = !Properties.Settings.Default.sortAscending;
            cbAscending.Checked = Properties.Settings.Default.sortAscending;

            chkRegulon.Checked = Properties.Settings.Default.regPlot;
            cbMapping.Checked = Properties.Settings.Default.tblMap;
            cbSummary.Checked = Properties.Settings.Default.tblSummary;
            cbCombined.Checked = Properties.Settings.Default.tblCombine;
            cbOperon.Checked = Properties.Settings.Default.tblOperon;

            cbClustered.Checked = Properties.Settings.Default.catPlot;
            cbDistribution.Checked = Properties.Settings.Default.distPlot;

            cbUseCategories.Checked = Properties.Settings.Default.useCat;
            cbUseRegulons.Checked = !Properties.Settings.Default.useCat;

            if (Properties.Settings.Default.categoryFile.Length == 0)
            {
                cbUseCategories.Checked = false; ;
                cbUseRegulons.Checked = true;
                Properties.Settings.Default.useCat = false;

            }

            cbUsePValues.Checked = Properties.Settings.Default.use_pvalues;
            cbUseFoldChanges.Checked = !Properties.Settings.Default.use_pvalues;
        }

        /// <summary>
        /// The initial load procedure of the Add-in. Initialize fields and labels depending on last known settings
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void GinRibbon_Load(object sender, RibbonUIEventArgs e)
        {

            LoadButtonStatus();

            if (Properties.Settings.Default.operonFile.Length == 0)
                btnOperonFile.Label = "No file selected";

            gAvailItems = PropertyItems("directionMapUnassigned");
            gUpItems = PropertyItems("directionMapUp");
            gDownItems = PropertyItems("directionMapDown");

            InitFields();

            PlotRoutines.theApp = gApplication;

            EnableOutputOptions(false);

            gExcelErrorValues = ((int[])Enum.GetValues(typeof(ExcelUtils.CVErrEnum))).ToList();

            btLoad.Enabled = System.IO.File.Exists(Properties.Settings.Default.referenceFile);

        }


        /// <summary>
        /// Enable/disable possible output buttons (i.e. table or charts).
        /// </summary>
        /// <param name="enable"></param>
        private void EnableOutputOptions(bool enable)
        {
            ebLow.Enabled = enable;
            ebMid.Enabled = enable;
            ebHigh.Enabled = enable;
            editMinPval.Enabled = enable;

            cbMapping.Enabled = enable;
            cbSummary.Enabled = enable;
            cbCombined.Enabled = enable;

            cbClustered.Enabled = enable;
            cbDistribution.Enabled = enable;
            chkRegulon.Enabled = enable;

            cbOrderFC.Enabled = enable;
            cbUseCategories.Enabled = enable && gCatOutput;
            cbUseRegulons.Enabled = enable && gCatOutput;

            cbOperon.Enabled = enable && gOperonOutput;

            cbUsePValues.Enabled = enable;
            cbUseFoldChanges.Enabled = enable;

            cbAscending.Enabled = enable;
            cbDescending.Enabled = enable;

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
        /// The main routine to map the data to a list of genes with their associated FC, p-values etc.. and a list of regulonss
        /// </summary>
        /// <param name="theCells"></param>
        /// <returns>A list of data genes</returns>
        private List<BsuRegulons> QueryResultTable(List<Excel.Range> theCells)
        {

            AddTask(TASKS.MAPPING_GENES_TO_REGULONS);

            // Copy the data from the selected cells as determined in the dialog earlier.
            object[,] rangeBSU = theCells[2].Value2;
            object[,] rangeFC = theCells[1].Value2;
            object[,] rangeP = theCells[0].Value2;


            // Initialize the list
            List<BsuRegulons> lList = new List<BsuRegulons>();

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

                // create a mapping entry
                BsuRegulons lMap = new BsuRegulons(lFC, lPvalue, lBSU);

                //  double check if BSU has a value 
                if (lMap.BSU.Length > 0)
                {
                    // find the entries that are linked by the same gene
                    SysData.DataRow[] results = Lookup(lMap.BSU);

                    // loop over the entries (=regulons) found
                    if (results.Length > 0)
                    {
                        string gene = results[0][Properties.Settings.Default.referenceGene].ToString();
                        lMap.GENE = gene;

                        // store every regulon and it's association if found in the user defined up/down regulation list

                        for (int r = 0; r < results.Length; r++)
                        {
                            string item = results[r][Properties.Settings.Default.referenceRegulon].ToString();
                            string direction = results[r][Properties.Settings.Default.referenceDIR].ToString();


                            if (item.Length > 0)
                            {
                                lMap.REGULONS.Add(item);

                                if (gUpItems.Contains(direction))
                                    lMap.UP.Add(r);

                                if (gDownItems.Contains(direction))
                                    lMap.DOWN.Add(r);
                            }
                        }
                    }
                }

                lList.Add(lMap);
            }

            RemoveTask(TASKS.MAPPING_GENES_TO_REGULONS);
            return lList;

        }

        /// <summary>
        /// This routine reformats the original data to a data table that can be exported to an excel sheet. The second data table contains the color formatting of the different cells.
        /// </summary>
        /// <param name="lResults"></param>
        /// <returns></returns>
        (SysData.DataTable, SysData.DataTable) PrepareResultTable(List<BsuRegulons> lResults)
        {
            SysData.DataTable myTable = new System.Data.DataTable("mytable");
            SysData.DataTable clrTable = new System.Data.DataTable("colortable");

            int maxcol = lResults[0].REGULONS.Count;

            // count max number of columns neccesary
            for (int r = 1; r < lResults.Count; r++)
                if (maxcol < lResults[r].REGULONS.Count)
                    maxcol = lResults[r].REGULONS.Count;

            // add BSU/gene/p-value/fc columns

            SysData.DataColumn bsuCol = new SysData.DataColumn("bsu", Type.GetType("System.String"));
            myTable.Columns.Add(bsuCol);
            SysData.DataColumn geneCol = new SysData.DataColumn("gene", Type.GetType("System.String"));
            myTable.Columns.Add(geneCol);
            SysData.DataColumn fcCol = new SysData.DataColumn("fc", Type.GetType("System.Double"));
            myTable.Columns.Add(fcCol);
            SysData.DataColumn pValCol = new SysData.DataColumn("pval", Type.GetType("System.Double"));
            myTable.Columns.Add(pValCol);

            // add count column
            SysData.DataColumn countCol = new SysData.DataColumn("count_col", Type.GetType("System.Int16"));
            myTable.Columns.Add(countCol);

            // add variable columns
            for (int c = 0; c < maxcol; c++)
            {
                SysData.DataColumn newCol = new SysData.DataColumn(string.Format("col_{0}", c + 1));
                myTable.Columns.Add(newCol);
                SysData.DataColumn clrCol = new SysData.DataColumn(string.Format("col_{0}", c + 1), Type.GetType("System.Int16"));
                clrTable.Columns.Add(clrCol);
            }

            // fill data
            for (int r = 0; r < lResults.Count; r++)
            {
                SysData.DataRow newRow = myTable.Rows.Add();

                newRow["bsu"] = lResults[r].BSU;
                newRow["gene"] = lResults[r].GENE;
                newRow["fc"] = lResults[r].FC;
                newRow["pval"] = lResults[r].PVALUE;

                newRow["count_col"] = lResults[r].TOT;
                SysData.DataRow clrRow = clrTable.Rows.Add();

                for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                    newRow[string.Format("col_{0}", c + 1)] = lResults[r].REGULONS[c];

                for (int c = 0; c < lResults[r].UP.Count; c++)
                    clrRow[lResults[r].UP[c]] = 1;

                for (int c = 0; c < lResults[r].DOWN.Count; c++)
                    clrRow[lResults[r].DOWN[c]] = -1;

            }

            return (myTable, clrTable);


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
        /// The main routine after mouse selection update. It reads in the data from the excel sheet to return Raw input in a list of BSU structures and a List of genes and their associated regulons 
        /// </summary>
        /// <param name="suppressOutput"></param>
        /// <returns></returns>        
        private (List<FC_BSU>, List<BsuRegulons>) GenerateOutput()
        {

            AddTask(TASKS.READ_SHEET_DATA);

            List<Excel.Range> theInputCells = new List<Excel.Range>()
            {
                gRangeP,
                gRangeFC,
                gRangeBSU
            };

            // set flag is data has changed       
            if (InputHasChanged() || gOutput == null || gList == null)
            {
                gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;
            }
            else
            {
                RemoveTask(TASKS.READ_SHEET_DATA);
                return (gOutput, gList);
            }

            int nrRows = gRangeP.Rows.Count;

            // generate the results for outputting the data and summary
            try
            {
                List<BsuRegulons> lResults = QueryResultTable(theInputCells);

                // copy the input to a List of structures
                List<FC_BSU> lOutput = new List<FC_BSU>();

                // for all genes register the positive or negative regulon association or none if not defined
                for (int r = 0; r < nrRows; r++)
                    for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                    {
                        int val = 0;
                        if (lResults[r].UP.Contains(c))
                            val = 1;
                        if (lResults[r].DOWN.Contains(c))
                            val = -1;

                        // augment data with regulon info
                        lOutput.Add(new FC_BSU(lResults[r].FC, lResults[r].REGULONS[c], val, lResults[r].PVALUE, lResults[r].GENE));
                    }

                RemoveTask(TASKS.READ_SHEET_DATA);

                return (lOutput, lResults);
            }
            catch
            {
                MessageBox.Show("Are you sure the columns do not contain text?");
                RemoveTask(TASKS.READ_SHEET_DATA);

                return (null, null);
            }

        }


        /// <summary>
        /// Create the worksheet that contains the basic mapping gene - regulon table
        /// </summary>
        /// <param name="bsuRegulons"></param>
        private void CreateMappingSheet(List<BsuRegulons> bsuRegulons)
        {
            (SysData.DataTable lTable, SysData.DataTable clrTbl) = PrepareResultTable(bsuRegulons);

            AddTask(TASKS.UPDATE_MAPPED_TABLE);

            int nrRows = lTable.Rows.Count;
            int startR = 2;
            int offsetColumn = 1;

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            RenameWorksheet(lNewSheet, "Mapped_");

            lNewSheet.Cells[1, 1] = "BSU";
            lNewSheet.Cells[1, 2] = "GENE";
            lNewSheet.Cells[1, 3] = "FC";
            lNewSheet.Cells[1, 4] = "PVALUE";
            lNewSheet.Cells[1, 5] = "TOT REGULONS";

            string lastColumn = lTable.Columns[lTable.Columns.Count - 1].ColumnName;
            lastColumn = lastColumn.Replace("col_", "");
            int maxreg = Int16.Parse(lastColumn);

            for (int i = 0; i < maxreg; i++)
                lNewSheet.Cells[1, i + 6] = string.Format("Regulon_{0}", i + 1);

            // copy data to excel sheet

            FastDtToExcel(lTable, lNewSheet, startR, offsetColumn, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);

            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[lTable.Rows.Count + 1, lTable.Columns.Count];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();

            // color cells according to table 
            ColorCells(clrTbl, lNewSheet, startR, offsetColumn + 5, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);

            RemoveTask(TASKS.UPDATE_MAPPED_TABLE);

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
        /// Create a summary of regulons and their positive and/or negative associated genes
        /// </summary>
        /// <param name="theTable"></param>
        private void CreateSummarySheet(SysData.DataTable theTable)
        {

            AddTask(TASKS.UPDATE_SUMMARY_TABLE);

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            RenameWorksheet(lNewSheet, "Summary_");

            int col = 1;


            Excel.Range top = lNewSheet.Cells[1, 4];
            Excel.Range bottom = lNewSheet.Cells[1, 11];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Observed Counts and directions";
            all.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            top = lNewSheet.Cells[1, 14];
            bottom = lNewSheet.Cells[1, 21];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Percentage";
            all.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            top = lNewSheet.Cells[1, 22];
            bottom = lNewSheet.Cells[1, 25];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Probability regulation direction";
            all.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            lNewSheet.Cells[2, col++] = "Regulon";
            lNewSheet.Cells[2, col++] = "Total number in database";
            lNewSheet.Cells[2, col++] = "Total number in dataset";

            lNewSheet.Cells[2, col++] = string.Format("UP >{0}", Properties.Settings.Default.fcHIGH);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >{1}", Properties.Settings.Default.fcHIGH, Properties.Settings.Default.fcMID);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >{1}", Properties.Settings.Default.fcMID, Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >0", Properties.Settings.Default.fcLOW);

            lNewSheet.Cells[2, col++] = string.Format("DOWN <0 & >=-{0}", Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <-{0} & >=-{1}", Properties.Settings.Default.fcMID, Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <=-{0} & >=-{1}", Properties.Settings.Default.fcHIGH, Properties.Settings.Default.fcMID);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <-{0}", Properties.Settings.Default.fcHIGH);

            lNewSheet.Cells[2, col++] = "Total Relevant";
            lNewSheet.Cells[2, col++] = "% Total Relevant";

            int colGreen = col;

            lNewSheet.Cells[2, col++] = string.Format("UP >{0}", Properties.Settings.Default.fcHIGH);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >{1}", Properties.Settings.Default.fcHIGH, Properties.Settings.Default.fcMID);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >{1}", Properties.Settings.Default.fcMID, Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >0", Properties.Settings.Default.fcLOW);

            lNewSheet.Cells[2, col++] = string.Format("DOWN <0 & >=-{0}", Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <-{0} & >=-{1}", Properties.Settings.Default.fcMID, Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <=-{0} & >=-{1}", Properties.Settings.Default.fcHIGH, Properties.Settings.Default.fcMID);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <-{0}", Properties.Settings.Default.fcHIGH);

            lNewSheet.Cells[2, col++] = "If activation";
            lNewSheet.Cells[2, col++] = "Support";
            lNewSheet.Cells[2, col++] = "If repression";
            lNewSheet.Cells[2, col++] = "Support";


            FastDtToExcel(theTable, lNewSheet, 3, 1, theTable.Rows.Count + 2, theTable.Columns.Count);

            // color the blocks of cells... not by direction but just to separate up from down regulated            
            top = lNewSheet.Cells[3, colGreen];
            bottom = lNewSheet.Cells[theTable.Rows.Count + 2, colGreen + 4];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            top = lNewSheet.Cells[3, colGreen + 4];
            bottom = lNewSheet.Cells[theTable.Rows.Count + 2, colGreen + 4 + 3];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);

            // set number formats for the different types of cells

            top = lNewSheet.Cells[3, 13];
            bottom = lNewSheet.Cells[2 + theTable.Rows.Count, 22];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.NumberFormat = "###%";


            top = lNewSheet.Cells[3, 24];
            bottom = lNewSheet.Cells[2 + theTable.Rows.Count, 24];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.NumberFormat = "###%";


            // fit the width of the columns
            top = lNewSheet.Cells[1, 1];
            bottom = lNewSheet.Cells[theTable.Rows.Count + 2, theTable.Columns.Count];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();


            RemoveTask(TASKS.UPDATE_SUMMARY_TABLE);

        }

        /// <summary>
        /// Reformat the 'raw' augmented data into a datatable that can be displayed on a worksheet
        /// </summary>
        /// <param name="aList"></param>
        /// <returns></returns>
        private SysData.DataTable ReformatResults(List<FC_BSU> aList)
        {
            // find unique regulons

            SysData.DataTable lTable = new SysData.DataTable("FC_BSU");
            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn geneColumn = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            SysData.DataColumn pvalColumn = new SysData.DataColumn("Pvalue", Type.GetType("System.Double"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            SysData.DataColumn dirColumn = new SysData.DataColumn("DIR", Type.GetType("System.Int32"));


            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(geneColumn);
            lTable.Columns.Add(fcColumn);
            lTable.Columns.Add(pvalColumn);
            lTable.Columns.Add(dirColumn);

            for (int r = 0; r < aList.Count; r++)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                lRow["Regulon"] = aList[r].BSU;
                lRow["FC"] = aList[r].FC;
                lRow["DIR"] = aList[r].DIR;
                lRow["Pvalue"] = aList[r].PVALUE;
                lRow["Gene"] = aList[r].GENE;
            }

            return lTable;

        }


        /// <summary>
        /// Calculate the nr of genes up-regulated and down-regulated by it's fold change and the expected direction based on the mode of the regulon
        /// </summary>
        /// <param name="aRow"></param>
        /// <returns>repressed up, activated up, repressed down, activated down, total</returns>
        private (int, int, int, int, int) CalculateFPRatio(SysData.DataRow[] aRow)
        {
            int nrActiveUP = 0, nrActiveDOWN = 0, nrRepressUP = 0, nrRepressDown = 0, nrTot = 0;

            // aRow from an FC_BSU table
            for (int i = 0; i < aRow.Length; i++)
            {
                double fcGene = (double)aRow[i]["FC"];
                int dirBSU = (int)aRow[i]["DIR"];
                double lowValue = Properties.Settings.Default.fcLOW;

                // if in repressed mode
                if (dirBSU < 0)
                {
                    // and fc < -lowValue
                    if (fcGene < -lowValue)
                    {
                        nrRepressUP += 1;
                        nrTot += 1;
                    }

                    // and fc > lowValue 
                    if (fcGene > lowValue)
                    {
                        nrRepressDown += 1;
                        nrTot += 1;
                    }
                }
                // if in activated mode
                if (dirBSU > 0)
                {
                    // and fc > lowValue
                    if (fcGene > lowValue)
                    {
                        nrActiveUP += 1;
                        nrTot += 1;
                    }

                    // fc < -lowValue
                    if (fcGene < -lowValue)
                    {
                        nrActiveDOWN += 1;
                        nrTot += 1;
                    }
                }

            }

            return (nrRepressUP, nrActiveUP, nrRepressDown, nrActiveDOWN, nrTot);
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
        /// Create an overview of the data mapped to their operons
        /// </summary>
        /// <param name="table"></param>
        private void CreateOperonSheet(SysData.DataTable table)
        {
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            RenameWorksheet(lNewSheet, "Operon_");


            int maxNrGenes = Int32.Parse(table.Compute("max([nrgenes])", string.Empty).ToString());

            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

            int col = 1;
            lNewSheet.Cells[1, col++] = "BSU";
            lNewSheet.Cells[1, col++] = "FC";
            lNewSheet.Cells[1, col++] = "P-Value";

            lNewSheet.Cells[1, col++] = "Gene";
            lNewSheet.Cells[1, col++] = "Operon Name";
            lNewSheet.Cells[1, col++] = "Nr operons";
            lNewSheet.Cells[1, col++] = "Nr genes";
            lNewSheet.Cells[1, col++] = "Operon";

            for (int c = 0; c < maxNrGenes; c++)
            {
                string colHeader = string.Format("FC Gene #{0}", c + 1);
                lNewSheet.Cells[1, c + col] = colHeader;
            }

            FastDtToExcel(table, lNewSheet, 2, 1, table.Rows.Count + 1, maxNrGenes + 4);


            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[table.Rows.Count + 1, maxNrGenes + 4];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();
            all.Rows.AutoFit();


            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;
        }


        /// <summary>
        ///  Create a combined table that combines the raw data and regulon summaries
        /// </summary>
        /// <param name="aUsageTbl"></param>
        /// <param name="lLst"></param>
        /// <returns></returns>
        private (SysData.DataTable, SysData.DataTable) CreateCombinedTable(SysData.DataTable aUsageTbl, List<BsuRegulons> lLst)
        {
            SysData.DataTable lTable = new SysData.DataTable();
            SysData.DataTable lColorTable = new SysData.DataTable();


            SysData.DataColumn col = new SysData.DataColumn("BSU", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("GENE", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            lTable.Columns.Add(col);


            col = new SysData.DataColumn("PVALUE", Type.GetType("System.Double"));
            lTable.Columns.Add(col);

            int maxRegulons = 0;
            for (int i = 0; i < lLst.Count; i++)
            {
                if (maxRegulons < lLst[i].REGULONS.Count)
                    maxRegulons = lLst[i].REGULONS.Count;
            }

            for (int i = 0; i < maxRegulons; i++)
            {
                col = new SysData.DataColumn(string.Format("Regulon_{0}", i + 1), Type.GetType("System.String"));
                lTable.Columns.Add(col);
                col = new SysData.DataColumn(string.Format("Regulon_{0}", i + 1), Type.GetType("System.Int16"));
                lColorTable.Columns.Add(col);

            }

            double lowVal = Properties.Settings.Default.fcLOW;

            // loop over all the genes found in the data 
            for (int r = 0; r < lLst.Count; r++)
            {
                // continue depending on value of lowest fc definition
                bool accept = Properties.Settings.Default.use_pvalues ? lLst[r].PVALUE < Properties.Settings.Default.pvalue_cutoff : Math.Abs(lLst[r].FC) > lowVal;

                if (true)
                {
                    SysData.DataRow lColorRow = lColorTable.Rows.Add();

                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["FC"] = lLst[r].FC;
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["GENE"] = lLst[r].GENE;
                    lRow["PVALUE"] = lLst[r].PVALUE;


                    double FC = lLst[r].FC;

                    for (int i = 0; i < lLst[r].REGULONS.Count; i++)
                    {

                        // check association direction 
                        bool posAssoc = lLst[r].UP.Contains(i);
                        bool negAssoc = lLst[r].DOWN.Contains(i);
                        // depending on the association in the table the cell color is red or green

                        int clrInt = posAssoc ? 1 : negAssoc ? -1 : 0;

                        SysData.DataRow[] lHit = aUsageTbl.Select(string.Format("Regulon = '{0}'", lLst[r].REGULONS[i]));
                        double nrUP = Double.Parse(lHit[0]["nr_UP"].ToString());
                        double nrDOWN = Double.Parse(lHit[0]["nr_DOWN"].ToString());
                        Double.TryParse(lHit[0]["perc_UP"].ToString(), out double percUP);
                        Double.TryParse(lHit[0]["perc_DOWN"].ToString(), out double percDOWN);

                        double percRel = Double.Parse(lHit[0]["totrelperc"].ToString());

                        string lVal = "";
                        string _down = "\u2193";
                        string _up = "\u2191";


                        // logical association
                        if ((posAssoc && FC > 0) || (negAssoc && FC < 0))
                        {
                            lVal = percUP.ToString("P0") + _up + percRel.ToString("P0") + "-tot";
                        }

                        if (nrUP == nrDOWN)
                            lVal = "0%-" + percRel.ToString("P0") + "-tot";

                        // false postive/negative
                        if ((posAssoc && FC < 0) || (negAssoc && FC > 0))
                        {
                            lVal = percDOWN.ToString("P0") + _down + percRel.ToString("P0") + "-tot";
                        }

                        lRow[string.Format("Regulon_{0}", i + 1)] = lLst[r].REGULONS[i] + " " + lVal;
                        lColorRow[string.Format("Regulon_{0}", i + 1)] = clrInt;
                    }
                }
            }


            for (int i = maxRegulons; i > 0; i--)
            {
                string columnName = string.Format("Regulon_{0}", i);
                object lRes = lTable.Compute(string.Format("COUNT({0})", columnName), "");
                int lCount = Int16.Parse(lRes.ToString());
                if (lCount == 0)
                {
                    lTable.Columns.Remove(columnName);
                    lColorTable.Columns.Remove(columnName);
                }
            }

            return (lTable, lColorTable);
        }

        /// <summary>
        /// Create a combined sheet in which not only the genes are mapped to regulons but also a summary of the associated regulons are displayed 
        /// </summary>
        /// <param name="aTable"></param>
        /// <param name="aClrTable"></param>
        private void CreateCombinedSheet(SysData.DataTable aTable, SysData.DataTable aClrTable)
        {

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            RenameWorksheet(lNewSheet, "Combined_");

            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

            int firstRow = 2;
            int firstCol = 1;
            int lastCol = aTable.Columns.Count + firstCol;
            int lastRow = aTable.Rows.Count + firstRow;

            Excel.Range top = lNewSheet.Cells[firstRow, firstCol];
            Excel.Range bottom = lNewSheet.Cells[lastRow, lastCol];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            int col = 1;

            lNewSheet.Cells[1, col++] = "BSU";
            lNewSheet.Cells[1, col++] = "GENE";
            lNewSheet.Cells[1, col++] = "FC";
            lNewSheet.Cells[1, col++] = "PVALUE";


            // determine the maximum number of regulons from the table that wass passed

            string lastColumn = aTable.Columns[aTable.Columns.Count - 1].ColumnName;
            lastColumn = lastColumn.Replace("Regulon_", "");
            int maxRegulons = Int16.Parse(lastColumn);

            for (int c = 0; c < maxRegulons; c++)
                lNewSheet.Cells[1, col++] = string.Format("Regulon_{0}", c + 1);

            FastDtToExcel(aTable, lNewSheet, 2, 1, aTable.Rows.Count + 1, aTable.Columns.Count);
            ColorCells(aClrTable, lNewSheet, 2, 5, aTable.Rows.Count + 1, aClrTable.Columns.Count);

            all.Columns.AutoFit();

            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;

        }

        /// <summary>
        /// Return a list of genes and their FCs that are linked by a single operon
        /// </summary>
        /// <param name="opid"></param>
        /// <param name="lLst"></param>
        /// <returns></returns>
        private (List<string>, List<double>) GetOperonGenesFC(/*string operon,*/ string opid, List<BsuRegulons> lLst)
        {

            SysData.DataRow[] lquery = gRefOperons.Select(string.Format("op_id = '{0}'", opid));

            List<string> _genes = new List<string>();
            List<double> _lfcs = new List<double>();

            foreach (DataRow row in lquery)
            {

                string lgene = row["gene"].ToString();
                _genes.Add(lgene);

                BsuRegulons result = lLst.Find(item => item.GENE == lgene);
                if (result != null)
                    _lfcs.Add(result.FC);
                else
                    _lfcs.Add(Double.NaN);
            }

            return (_genes, _lfcs);
        }

        /// <summary>
        /// Create the data table that can be used to display the data augmented with operon info in an Excel sheet.
        /// </summary>
        /// <param name="aUsageTbl"></param>
        /// <param name="lLst"></param>
        /// <returns></returns>
        private SysData.DataTable CreateOperonTable(List<BsuRegulons> lLst)
        {

            SysData.DataTable lTable = new SysData.DataTable();

            SysData.DataColumn col = new SysData.DataColumn("BSU", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("P-value", Type.GetType("System.Double"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("operon", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("nroperons", Type.GetType("System.Int16"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("nrgenes", Type.GetType("System.Int16"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("operon_genes", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            for (int nr = 0; nr < maxGenesPerOperon; nr++)
            {
                col = new SysData.DataColumn(string.Format("gene_{0}", nr + 1), Type.GetType("System.Double"));
                lTable.Columns.Add(col);
            }


            //double lowVal = Properties.Settings.Default.fcLOW;

            // loop over the all the genes found
            for (int r = 0; r < lLst.Count; r++)
            {

                string geneName = lLst[r].GENE;
                //double lFC = lLst[r].FC;
                //double lPval = lLst[r].PVALUE;

                List<string> luOperons = new List<string>();


                // possibly multiple operons for a single gene
                SysData.DataRow[] lOperons = gRefOperons.Select(string.Format("gene='{0}'", geneName));


                string operon = "";

                // create a list of 'other' genes
                List<string> lgenes = new List<string>();

                // create a list of 'other' genes' FCs
                List<double> lFCs = new List<double>();

                //string opgenes = "";
                List<List<string>> llgenes = new List<List<string>>();

                int _m = 0;
                int _maxm = _m;
                // loop over all operons first to determine leading operon .. i.e. the one with the most number of genes associated
                foreach (DataRow row in lOperons)
                {

                    luOperons.Add(row["operon"].ToString());
                    // an operon in return can have multiple genes associated with it.. register it and get the FCs
                    (List<string> _lgenes, List<double> _lFCs) = GetOperonGenesFC(row["op_id"].ToString(), lLst);

                    // add to list
                    llgenes.Add(_lgenes);

                    // if newly found genes is larger than previously found, store the associated values
                    if (_lgenes.Count > lgenes.Count)
                    {
                        _maxm = _m;
                        operon = row["operon"].ToString();
                        lgenes = new List<string>(_lgenes);
                        lFCs = new List<double>(_lFCs);
                    }

                    _m++;

                }

                // count nr of operons
                int noperons = luOperons.Count();

                // if any operon is found
                if (operon.Length > 0)
                {
                    // assign operon with most genes as leading operon
                    operon = luOperons[_maxm];
                    lgenes = llgenes[_maxm];
                    string opgenes = string.Join("-", lgenes.ToArray());
                    llgenes.Remove(lgenes);

                    // combine multiple orperons in a hyphenated string
                    foreach (List<string> _item in llgenes)
                    {
                        opgenes = opgenes + Environment.NewLine + string.Join("-", _item.ToArray());
                    }

                    int nrgenes = lgenes.Count;

                    // add a row
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["FC"] = lLst[r].FC;
                    lRow["P-Value"] = lLst[r].PVALUE;

                    lRow["gene"] = geneName;
                    lRow["operon"] = operon;
                    lRow["nroperons"] = noperons;
                    lRow["nrgenes"] = nrgenes;
                    lRow["operon_genes"] = opgenes;

                    // copy the FCs

                    for (int i = 0; i < nrgenes; i++)
                    {
                        if (!(lFCs[i] is Double.NaN))
                            lRow[string.Format("gene_{0}", i + 1)] = lFCs[i];
                    }
                }
                else // there's no linked operon found for this entry
                {

                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["FC"] = lLst[r].FC;
                    lRow["P-Value"] = lLst[r].PVALUE;

                }

            }
            return lTable;
        }



        /// <summary>
        /// Create the tables for the summary and combined info sheets
        /// </summary>
        /// <param name="aList"></param>
        /// <returns></returns>
        private (SysData.DataTable, SysData.DataTable) CreateUsageTable(List<FC_BSU> aList)
        {
            {
                SysData.DataTable _fc_BSU = ReformatResults(aList);

                SysData.DataTable lTable = new SysData.DataTable();
                SysData.DataTable lTableCombine = new SysData.DataTable(); // table for combined summary

                string _down = "\u2193";
                string _up = "\u2191";

                double lFClow = Properties.Settings.Default.fcLOW;
                double lFCmid = Properties.Settings.Default.fcMID;
                double lFChigh = Properties.Settings.Default.fcHIGH;

                // find number of unique regulons
                HashSet<string> lUnique = new HashSet<string>();

                for (int r = 0; r < aList.Count; r++)
                    lUnique.Add(aList[r].BSU);

                // add the columns per defined FC range
                SysData.DataColumn col = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
                lTable.Columns.Add(col);


                col = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
                lTable.Columns.Add(col);
                col = new SysData.DataColumn("CountData", Type.GetType("System.Int16"));
                lTable.Columns.Add(col);


                for (int i = 3; i >= 0; i--)
                {
                    col = new SysData.DataColumn(string.Format("up{0}", i + 1), Type.GetType("System.Double"));
                    lTable.Columns.Add(col);
                }

                for (int i = 0; i < 4; i++)
                {
                    col = new SysData.DataColumn(string.Format("down{0}", i + 1), Type.GetType("System.Double"));
                    lTable.Columns.Add(col);
                }

                col = new SysData.DataColumn("totrel", Type.GetType("System.Int16"));
                lTable.Columns.Add(col);

                col = new SysData.DataColumn("percrel", Type.GetType("System.Double"));
                lTable.Columns.Add(col);


                for (int i = 3; i >= 0; i--)
                {
                    col = new SysData.DataColumn(string.Format("perc_up{0}", i + 1), Type.GetType("System.Double"));
                    lTable.Columns.Add(col);
                }

                for (int i = 0; i < 4; i++)
                {
                    col = new SysData.DataColumn(string.Format("perc_down{0}", i + 1), Type.GetType("System.Double"));
                    lTable.Columns.Add(col);
                }


                // if activation
                col = new SysData.DataColumn("perc_UP", Type.GetType("System.Double"));
                lTable.Columns.Add(col);


                col = new SysData.DataColumn("activated", Type.GetType("System.String"));
                lTable.Columns.Add(col);

                // if repression

                col = new SysData.DataColumn("perc_DOWN", Type.GetType("System.Double"));
                lTable.Columns.Add(col);


                col = new SysData.DataColumn("repressed", Type.GetType("System.String"));
                lTable.Columns.Add(col);


                /* define the combined table */

                col = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
                lTableCombine.Columns.Add(col);

                if (gOperonOutput)
                {
                    col = new SysData.DataColumn("operon", Type.GetType("System.String"));
                    lTable.Columns.Add(col);
                }

                col = new SysData.DataColumn("totrelperc", Type.GetType("System.Double"));
                lTableCombine.Columns.Add(col);

                col = new SysData.DataColumn("perc_DOWN", Type.GetType("System.Double"));
                lTableCombine.Columns.Add(col);
                col = new SysData.DataColumn("perc_UP", Type.GetType("System.Double"));
                lTableCombine.Columns.Add(col);


                col = new SysData.DataColumn("nr_DOWN", Type.GetType("System.Int16"));
                lTableCombine.Columns.Add(col);
                col = new SysData.DataColumn("nr_UP", Type.GetType("System.Int16"));
                lTableCombine.Columns.Add(col);



                // file the table
                foreach (string reg in lUnique)
                {
                    int up1 = 0;
                    int up2 = 0;
                    int up3 = 0;
                    int up4 = 0;
                    int down1 = 0;
                    int down2 = 0;
                    int down3 = 0;
                    int down4 = 0;

                    // lookup regulon in global statistic table
                    SysData.DataRow[] _tmp2 = gRefStats.Select(string.Format("Regulon='{0}'", reg));

                    // calculate usage statistics in dataset
                    SysData.DataRow[] _tmp = _fc_BSU.Select(string.Format("Regulon='{0}'", reg));

                    (int repressedUP, int activatedUP, int repressedDOWN, int activatedDOWN, int nrTOT) = CalculateFPRatio(_tmp);

                    // up1-up4, down1-down4 contain the observed regulations of the genes with a specific fc

                    for (int _r = 0; _r < _tmp.Length; _r++)
                    {
                        double fc = (double)_tmp[_r]["FC"];
                        if (fc > 0 & fc <= lFClow)
                            up1 += 1;
                        if (fc > lFClow & fc <= lFCmid)
                            up2 += 1;
                        if (fc > lFCmid & fc <= lFChigh)
                            up3 += 1;
                        if (fc > lFChigh)
                            up4 += 1;

                        if (fc < 0 & fc >= -lFClow)
                            down1 += 1;
                        if (fc < -lFClow & fc >= -lFCmid)
                            down2 += 1;
                        if (fc < -lFCmid & fc >= -lFChigh)
                            down3 += 1;
                        if (fc < -lFChigh)
                            down4 += 1;

                    }

                    // add a new row to the main table

                    SysData.DataRow lNewRow = lTable.Rows.Add();

                    lNewRow["CountData"] = _tmp.Length;
                    int _lcount = Int16.Parse(_tmp2[0]["Count"].ToString());

                    lNewRow["Count"] = _lcount;


                    lNewRow["Regulon"] = reg;
                    lNewRow["down1"] = down1;
                    lNewRow["down2"] = down2;
                    lNewRow["down3"] = down3;
                    lNewRow["down4"] = down4;
                    lNewRow["up1"] = up1;
                    lNewRow["up2"] = up2;
                    lNewRow["up3"] = up3;
                    lNewRow["up4"] = up4;

                    lNewRow["perc_up1"] = (double)up1 / (double)_tmp.Length;
                    lNewRow["perc_up2"] = (double)up2 / (double)_tmp.Length;
                    lNewRow["perc_up3"] = (double)up3 / (double)_tmp.Length;
                    lNewRow["perc_up4"] = (double)up4 / (double)_tmp.Length;
                    lNewRow["perc_down1"] = (double)down1 / (double)_tmp.Length;
                    lNewRow["perc_down2"] = (double)down2 / (double)_tmp.Length;
                    lNewRow["perc_down3"] = (double)down3 / (double)_tmp.Length;
                    lNewRow["perc_down4"] = (double)down4 / (double)_tmp.Length;

                    // skip up1 and down1 .. they do not fall in the relevant class
                    lNewRow["totrel"] = up2 + up3 + up4 + down2 + down3 + down4; // nrTOT;
                    lNewRow["percrel"] = (double)(up2 + up3 + up4 + down2 + down3 + down4) / (double)_lcount; // nrTOT;

                    // nrUP and nrDOWN contain the counts of those genes that were defined as up or down regulated that had a 'significant' fc.
                    // this was, false positive can be identified
                    if (nrTOT > 0)
                    {
                        lNewRow["perc_DOWN"] = (double)(repressedDOWN + activatedDOWN) / (double)(nrTOT);
                        lNewRow["perc_UP"] = (double)(repressedUP + activatedUP) / (double)(nrTOT);
                    }

                    lNewRow["activated"] = repressedUP.ToString() + _down + activatedUP.ToString() + _up;
                    lNewRow["repressed"] = activatedDOWN.ToString() + _down + repressedDOWN.ToString() + _up;

                    // add a new row to the combined table

                    lNewRow = lTableCombine.Rows.Add();

                    // skip the up1 and down1 values, they don't make the cut.
                    double lRat = 0;
                    if (int.TryParse(_tmp2[0]["Count"].ToString(), out int totcount))
                    {
                        int totrel = up2 + up3 + up4 + down2 + down3 + down4;
                        lRat = (double)totrel / (double)totcount;

                    }

                    lNewRow["totrelperc"] = lRat;
                    lNewRow["Regulon"] = reg;

                    if (nrTOT > 0)
                    {
                        lNewRow["perc_DOWN"] = (double)(repressedDOWN + activatedDOWN) / (double)(nrTOT);
                        lNewRow["perc_UP"] = (double)(repressedUP + activatedUP) / (double)(nrTOT);
                    }

                    lNewRow["nr_DOWN"] = repressedDOWN + activatedDOWN;
                    lNewRow["nr_UP"] = repressedUP + activatedUP;

                }

                SysData.DataView dv = lTable.DefaultView;
                dv.Sort = "totrel desc";

                return (dv.ToTable(), lTableCombine);
            }
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
        /// Load the user-defined up/down definitions
        /// </summary>
        private void LoadDirectionOptions()
        {
            SysData.DataView view = new SysData.DataView(gRefWB);
            SysData.DataTable distinctValues = view.ToTable(true, Properties.Settings.Default.referenceDIR);

            foreach (SysData.DataRow row in distinctValues.Rows)
            {
                gAvailItems.Add(row.ItemArray[0].ToString());
            }
        }

        /// <summary>
        /// Load the regulon data  
        /// </summary>

        private void LoadWorksheets()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.referenceFile);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.referenceSheetName = ws.Name;


            excelworkBook.Close();

            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }


        /// <summary>
        /// Load the operon data
        /// </summary>
        private void LoadOperonSheet()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.operonFile);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.operonSheet = ws.Name;


            excelworkBook.Close();

            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }

        /// <summary>
        /// Load the operon data
        /// </summary>
        private void LoadCatFile()
        {
            Microsoft.Office.Interop.Excel.Application excel = (Microsoft.Office.Interop.Excel.Application)Globals.ThisAddIn.Application;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;

            Excel.Workbook excelworkBook = excel.Workbooks.Open(Properties.Settings.Default.categoryFile);
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.catSheet = ws.Name;


            excelworkBook.Close();

            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }

        /// <summary>
        /// Fill the dropdown boxes and select the last known (stored) selected value
        /// </summary>
        private void Fill_DropDownBoxes()
        {
            gApplication.EnableEvents = false;

            ddBSU.Items.Clear();
            ddRegulon.Items.Clear();
            ddGene.Items.Clear();
            ddDir.Items.Clear();

            foreach (string s in gColNames)
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

            ddItem = GetItemByValue(ddRegulon, Properties.Settings.Default.referenceRegulon);
            if (ddItem != null)
                ddRegulon.SelectedItem = ddItem;

            ddItem = GetItemByValue(ddDir, Properties.Settings.Default.referenceDIR);
            if (ddItem != null)
                ddDir.SelectedItem = ddItem;

            ddItem = GetItemByValue(ddGene, Properties.Settings.Default.referenceGene);
            if (ddItem != null)
                ddGene.SelectedItem = ddItem;

            ddBSU.Enabled = true;
            ddRegulon.Enabled = true;
            ddDir.Enabled = true;
            ddGene.Enabled = true;
            btRegDirMap.Enabled = true;

            gApplication.EnableEvents = true;


        }

        /// <summary>
        /// Reset main variables to initial values
        /// </summary>
        private void ResetTables()
        {
            gOutput = null;
            gList = null;
            gSummary = null;
            //gInputRange = null;
            gOldRangeBSU = "";
            gOldRangeFC = "";
            gOldRangeP = "";
            EnableOutputOptions(false);
            btApply.Enabled = false;
            btPlot.Enabled = false;

        }


        /// <summary>
        /// Load FC values from last known settings
        /// </summary>

        private void LoadFCDefaults()
        {
            ebLow.Text = Properties.Settings.Default.fcLOW.ToString();
            ebMid.Text = Properties.Settings.Default.fcMID.ToString();
            ebHigh.Text = Properties.Settings.Default.fcHIGH.ToString();
            editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();
        }


        //private void EnableItems(bool enable)
        //{
        //    btLoad.Enabled = enable;
        //    ddBSU.Enabled = enable;
        //    ddRegulon.Enabled = enable;
        //    ddGene.Enabled = enable;
        //    //btPlot.Enabled = enable;
        //    //edtMaxGroups.Enabled = enable;
        //    //btnPalette.Enabled = enable;

        //}


        /// <summary>
        /// Flag that defintion for BSU (=regulon code) has changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void DropDown_BSU_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceBSU = ddBSU.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
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
        }

        /// <summary>
        /// The definitions for up/down regulations have been changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_RegulonDirectionMap_Click(object sender, RibbonControlEventArgs e)
        {
            dlgUpDown dlgUD = new dlgUpDown(gAvailItems, gUpItems, gDownItems);
            dlgUD.ShowDialog();

            StoreValue("directionMapUnassigned", gAvailItems);
            StoreValue("directionMapUp", gUpItems);
            StoreValue("directionMapDown", gDownItems);

            SetFlags(UPDATE_FLAGS.ALL);

        }

        /// <summary>
        /// The column mapping identifier has changed. All data needs to be updated.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DropDown_RegulonDirection_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceDIR = ddDir.SelectedItem.Label;
            gAvailItems.Clear();
            gUpItems.Clear();
            gDownItems.Clear();
            LoadDirectionOptions();
            SetFlags(UPDATE_FLAGS.ALL);
        }

        /// <summary>
        /// Routine to check if changes made to textbox are ok, if not reset to previous value.
        /// </summary>
        /// <param name="bx"></param>

        private void ValidateTextBoxData(RibbonEditBox bx)
        {

            bool low = false;
            bool mid = false;
            bool high = false;

            if (bx.Equals(ebLow))
                low = true;
            if (bx.Equals(ebMid))
                mid = true;
            if (bx.Equals(ebHigh))
                high = true;

            // can still add range checks e.g. high > mid > low  

            if (double.TryParse(bx.Text, out double val))
            {
                // set the text value to what is parsed
                bx.Text = val.ToString();
                if (low)
                    Properties.Settings.Default.fcLOW = val;
                if (mid)
                    Properties.Settings.Default.fcMID = val;
                if (high)
                    Properties.Settings.Default.fcHIGH = val;

                SetFlags(UPDATE_FLAGS.FC_dependent);
            }
            else
            {
                if (low)
                    ebLow.Text = Properties.Settings.Default.fcLOW.ToString();
                if (mid)
                    ebMid.Text = Properties.Settings.Default.fcMID.ToString();
                if (high)
                    ebHigh.Text = Properties.Settings.Default.fcHIGH.ToString();
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
        /// Check changes in textbox Mid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_Mid_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ValidateTextBoxData(ebMid);
        }

        /// <summary>
        /// Check changes in textbox High
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void TextBox_High_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ValidateTextBoxData(ebHigh);
        }

        /// <summary>
        /// The main routine after the load button has been selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Load_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;
            if (LoadData())
            {
                gOperonOutput = LoadOperonData();
                gCatOutput = LoadCategoryData();

                Fill_DropDownBoxes();
                if (gDownItems.Count == 0 && gUpItems.Count == 0 && gAvailItems.Count == 0)
                    LoadDirectionOptions();

                btnSelect.Enabled = true;
                toggleButton1.Enabled = true;
                LoadFCDefaults();
                ResetTables();

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
            if (!(Properties.Settings.Default.catPlot || Properties.Settings.Default.regPlot || Properties.Settings.Default.distPlot))
            {
                MessageBox.Show("Please select at least one plot to generate");
                return;
            }


            if ((Properties.Settings.Default.catPlot || Properties.Settings.Default.regPlot)) //& gNeedsUpdate.Check(UPDATE_FLAGS.PCat))
            {
                dlgTreeView dlg = new dlgTreeView(categoryView: cbUseCategories.Checked, spreadingOptions: Properties.Settings.Default.catPlot, rankingOptions: Properties.Settings.Default.regPlot);

                if (gCategories != null && cbUseCategories.Checked)
                {
                    dlg.populateTree(gCategories);
                }
                if (gRefWB != null && !cbUseCategories.Checked)
                {
                    dlg.populateTree(gRefWB, cat: false);
                }

                if (dlg.ShowDialog() == DialogResult.OK)
                {

                    if ((gOutput == null || gSummary == null) || gNeedsUpdate.Check(UPDATE_FLAGS.TMapped))
                    {
                        (gOutput, gList) = GenerateOutput();

                        if (gOutput != null && gList != null)
                        {
                            UnSetFlags(UPDATE_FLAGS.TMapped);
                            (gSummary, gCombineInfo) = CreateUsageTable(gOutput);
                            UnSetFlags(UPDATE_FLAGS.TCombined);
                        }
                    }


                    if ((gOutput != null && gSummary != null && dlg.GetSelection().Count() > 0))
                    {
                        if (Properties.Settings.Default.catPlot)
                        {
                            SpreadingPlot(gOutput, gSummary, dlg.GetSelection(), topTenFC: dlg.getTopFC(), topTenP: dlg.getTopP(), outputTable: dlg.selectTableOutput());

                        }

                        if (Properties.Settings.Default.regPlot && gList != null)
                        {
                            RankingPlot(gOutput, gSummary, dlg.GetSelection());
                        }
                    }
                }

            }


            if (Properties.Settings.Default.distPlot)
            {
                if (gOutput == null || gSummary == null || gNeedsUpdate.Check(UPDATE_FLAGS.TMapped))
                {
                    (gOutput, gList) = GenerateOutput();
                    if (gOutput != null && gList != null)
                    {
                        UnSetFlags(UPDATE_FLAGS.TMapped);
                        (gSummary, gCombineInfo) = CreateUsageTable(gOutput);
                        UnSetFlags(UPDATE_FLAGS.TCombined);
                    }
                }

                if (gOutput != null)
                {
                    DistributionPlot(gOutput);
                }
            }
        }

        /// <summary>
        /// The routine after apply has been selected to created the worksheets.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void Button_Apply_Click(object sender, RibbonControlEventArgs e)
        {


            if (!(Properties.Settings.Default.tblMap || Properties.Settings.Default.tblSummary || Properties.Settings.Default.tblCombine || Properties.Settings.Default.tblOperon))
            {
                MessageBox.Show("Please select at least one output table to generate");
                return;
            }

            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            if (NoUpdate())
                return;

            if (Properties.Settings.Default.tblMap)
                CreateMappingSheet(gList);

            if ((gSummary == null && gCombineInfo == null) || NeedsUpdate(UPDATE_FLAGS.TSummary))
            {
                (gSummary, gCombineInfo) = CreateUsageTable(gOutput);
                UnSetFlags(UPDATE_FLAGS.TSummary);
            }

            if (gSummary != null && Properties.Settings.Default.tblSummary)
                CreateSummarySheet(gSummary);


            if (Properties.Settings.Default.tblCombine) // can combine table/sheet because it's a quick routine
            {
                (SysData.DataTable lCombined, SysData.DataTable lClrTable) = CreateCombinedTable(gCombineInfo, gList);
                CreateCombinedSheet(lCombined, lClrTable);
                UnSetFlags(UPDATE_FLAGS.TCombined);

            }


            if (Properties.Settings.Default.tblOperon && gOperonOutput) // can combine table/sheet because it's a quick routine
            {
                SysData.DataTable tblOperon = CreateOperonTable(gList);
                CreateOperonSheet(tblOperon);
                UnSetFlags(UPDATE_FLAGS.TOperon);
            }

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
        private element_fc CatElements2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements, int topTenFC = -1, int topTenP = -1)
        {

            //List<element_fc> element_Fcs = new List<element_fc>();
            element_fc element_Fcs = new element_fc();
            SysData.DataView dataViewCat = gCategories.AsDataView();


            List<summaryInfo> _All = new List<summaryInfo>();
            List<summaryInfo> _Pos = new List<summaryInfo>();
            List<summaryInfo> _Neg = new List<summaryInfo>();

            foreach (cat_elements ce in cat_Elements)
            {
                string categories = string.Join(",", ce.elements.ToArray());
                categories = string.Join(",", categories.Split(',').Select(x => $"'{x}'"));
                dataViewCat.RowFilter = String.Format("catid_short in ({0})", categories);

                HashSet<string> genes = new HashSet<string>();
                foreach (DataRow _row in dataViewCat.ToTable().Rows)
                {
                    genes.Add(_row[gCategoryGeneColumn].ToString());
                }

                string genesFormat = string.Join(",", genes.ToArray());
                genesFormat = string.Join(",", genesFormat.Split(',').Select(x => $"'{x}'"));
                dataView.RowFilter = String.Format("Gene in ({0})", genesFormat);

                SysData.DataTable _dt = dataView.ToTable(true, "Gene", "FC", "Pvalue");

                summaryInfo __All = new summaryInfo();
                summaryInfo __Pos = new summaryInfo();
                summaryInfo __Neg = new summaryInfo();

                __All.catName = string.Format("{0} ({1})", ce.catName, _dt.Rows.Count);
                __Pos.catName = string.Format("{0} ({1})", ce.catName, _dt.Rows.Count);
                __Neg.catName = string.Format("{0} ({1})", ce.catName, _dt.Rows.Count);


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

                        _genesA.Add(_dt.Rows[i]["Gene"].ToString());
                        _fcsA.Add(fc);
                        _pvaluesA.Add(double.Parse(_dt.Rows[i]["Pvalue"].ToString()));

                        if (fc >= 0)
                        {
                            _genesP.Add(_dt.Rows[i]["Gene"].ToString());
                            _fcsP.Add(fc);
                            _pvaluesP.Add(double.Parse(_dt.Rows[i]["Pvalue"].ToString()));
                        }
                        else if (fc < 0)
                        {
                            _genesN.Add(_dt.Rows[i]["Gene"].ToString());
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


                __Pos.p_average = _pvaluesP.Count > 0 ? _pvaluesP.paverage() : Double.NaN;
                __Neg.p_average = _pvaluesN.Count > 0 ? _pvaluesN.paverage() : Double.NaN;
                __All.p_average = _pvaluesA.Count > 0 ? _pvaluesA.paverage() : Double.NaN;

                __Pos.p_mad = _pvaluesP.Count > 0 ? _pvaluesP.mad() : Double.NaN;
                __Neg.p_mad = _pvaluesN.Count > 0 ? _pvaluesN.mad() : Double.NaN;
                __All.p_mad = _pvaluesA.Count > 0 ? _pvaluesA.mad() : Double.NaN;

                _All.Add(__All);
                _Pos.Add(__Pos);
                _Neg.Add(__Neg);
            }

            element_Fcs.All = _All;
            element_Fcs.Activated = _Pos;
            element_Fcs.Repressed = _Neg;

            // sort the values according to average FC and order in the preferred direction
            if (Properties.Settings.Default.useSort)
            {
                double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();

                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList();

                if (Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if (topTenFC > 0) // if top FC is selected only select the top X values using absolute FC values. The default order is descending
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
            else if (topTenP > 0) // do the same if top X p-value is selected. Because we are using -log(p) value the default order is descending
            {
                // only useful if top N selected is smaller then total number of items
                if (topTenP < element_Fcs.All.Count)
                {
                    // assertion... it's -10 log(p) -> so the higher, the better
                    double[] __values = element_Fcs.All.Select(x => -Math.Log(x.p_average)).ToArray();
                    var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();
                    List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                    element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).Where(x => x.fc_values.Length > 0).ToList();
                    element_Fcs.All = element_Fcs.All.GetRange(0, topTenP);
                }
                // if selected, reverse the order
                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }

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

        private element_fc Regulons2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements, int topTenFC = -1, int topTenP = -1)
        {
            element_fc element_Fcs = new element_fc();

            List<summaryInfo> _All = new List<summaryInfo>();
            List<summaryInfo> _Act = new List<summaryInfo>();
            List<summaryInfo> _Rep = new List<summaryInfo>();

            foreach (cat_elements el in cat_Elements)
            {
                dataView.RowFilter = String.Format("Regulon='{0}'", el.catName);

                SysData.DataTable _dataTable = dataView.ToTable();

                // find genes for the regulon/category

                summaryInfo __All = new summaryInfo();
                summaryInfo __Act = new summaryInfo();
                summaryInfo __Rep = new summaryInfo();

                __All.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                __Act.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                __Rep.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);

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
                        string _geneName = _dataTable.Rows[i]["Gene"].ToString();
                        _genesT.Add(_geneName);
                        _fcsT.Add(fc);
                        _pvaluesT.Add(double.Parse(_dataTable.Rows[i]["Pvalue"].ToString()));
                    }


                    DataRow[] _inhibited = _dataTable.Select("(FC<0 AND DIR>0) OR (FC>0 AND DIR<0) ");
                    for (int i = 0; i < _inhibited.Length; i++)
                    {
                        double fc = (double)_inhibited[i]["FC"];
                        string _geneName = _inhibited[i]["Gene"].ToString();
                        _genesR.Add(_geneName);
                        _fcsR.Add(fc);
                        _pvaluesR.Add(double.Parse(_inhibited[i]["Pvalue"].ToString()));
                    }


                    DataRow[] _activated = _dataTable.Select("(FC>0 AND DIR>0) OR (FC<0 AND DIR<0) ");
                    for (int i = 0; i < _activated.Length; i++)
                    {
                        double fc = (double)_activated[i]["FC"];
                        string _geneName = _activated[i]["Gene"].ToString();
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

                __Act.p_average = _pvaluesA.Count > 0 ? _pvaluesA.paverage() : Double.NaN;
                __Rep.p_average = _pvaluesR.Count > 0 ? _pvaluesR.paverage() : Double.NaN;
                __All.p_average = _pvaluesT.Count > 0 ? _pvaluesT.paverage() : Double.NaN;

                __Act.p_mad = _pvaluesA.Count > 0 ? _pvaluesA.mad() : Double.NaN;
                __Rep.p_mad = _pvaluesR.Count > 0 ? _pvaluesR.mad() : Double.NaN;
                __All.p_mad = _pvaluesT.Count > 0 ? _pvaluesT.mad() : Double.NaN;

                _All.Add(__All);
                _Act.Add(__Act);
                _Rep.Add(__Rep);


                element_Fcs.All = _All;
                element_Fcs.Activated = _Act;
                element_Fcs.Repressed = _Rep;

            }


            //default sort option is by average FC.
            if (Properties.Settings.Default.useSort)
            {
                double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();

                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList();

                if (Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if (topTenFC > 0) // top X FC is based on abs average FC.
            {
                // only useful if top N selected is smaller then total number of items
                if (topTenFC < element_Fcs.All.Count)
                {
                    double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                    var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                    List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                    // remove elements with no genes associated
                    element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList().Where(x => x.fc_values.Length > 0).ToList();
                    element_Fcs.All = element_Fcs.All.GetRange(0, topTenFC);
                }
                // reverse if selected
                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if (topTenP > 0) // top X p-value is based on descending, because of -log10(p) transformation
            {
                // only useful if top N selected is smaller then total number of items
                if (topTenP < element_Fcs.All.Count)
                {
                    // assertion... it's -10 log(p) -> so the higher, the better
                    double[] __values = element_Fcs.All.Select(x => -Math.Log(x.p_average)).ToArray();
                    var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();
                    List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                    // remove elements with no genes associated
                    element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList().Where(x => x.fc_values.Length > 0).ToList();
                    element_Fcs.All = element_Fcs.All.GetRange(0, topTenP);
                }

                // reverse if selected
                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }


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
        /// Create the distribution plot
        /// </summary>
        /// <param name="aOutput"></param>        
        private void DistributionPlot(List<FC_BSU> aOutput)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            SysData.DataTable _fc_BSU_ = ReformatResults(aOutput);
            SysData.DataTable _fc_BSU = GetDistinctRecords(_fc_BSU_, new string[] { "Gene", "FC" });

            (List<double> sFC, List<int> sIdx) = SortedFoldChanges(_fc_BSU);

            int chartNr = NextWorksheet("DistributionPlot_");
            string chartName = "DistributionPlot_" + chartNr.ToString();

            PlotRoutines.CreateDistributionPlot(sFC, sIdx, chartName);
            this.RibbonUI.ActivateTab("TabGINtool");


            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }

        /// <summary>
        /// Create the spreading plot that indicates the genes and FCs associated with a category or regulon. Optionally also output to a worksheet
        /// </summary>
        /// <param name="aOutput"></param>
        /// <param name="aSummary"></param>
        /// <param name="cat_Elements"></param>
        /// <param name="topTenFC"></param>
        /// <param name="topTenP"></param>
        /// <param name="outputTable"></param>
        private void SpreadingPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements, int topTenFC = -1, int topTenP = -1, bool outputTable = false)
        {

            AddTask(TASKS.CATEGORY_CHART);

            SysData.DataTable _fc_BSU = ReformatResults(aOutput);
            cat_Elements = GetUniqueElements(cat_Elements);

            // HashSet ensures unique list
            HashSet<string> lRegulons = new HashSet<string>();

            foreach (SysData.DataRow row in aSummary.Rows)
                lRegulons.Add(row.ItemArray[0].ToString());

            SysData.DataView dataView = _fc_BSU.AsDataView();
            element_fc catPlotData;
            if (Properties.Settings.Default.useCat)
            {
                catPlotData = CatElements2ElementsFC(dataView, cat_Elements, topTenFC, topTenP);
            }
            else
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements, topTenFC: topTenFC, topTenP: topTenP); // need to alter caller

            string postFix = topTenFC > -1 ? string.Format("Top{0}FC", topTenFC) : (topTenP > -1 ? string.Format("Top{0}P", topTenP) : "");
            string chartBase = (Properties.Settings.Default.useCat ? string.Format("CatSpreadPlot{0}_", postFix) : string.Format("RegSpreadPlot{0}_", postFix));
            int chartNr = NextWorksheet(chartBase);
            string chartName = chartBase + chartNr.ToString();
            PlotRoutines.CreateCategoryPlot(catPlotData, chartName);

            if (outputTable)
            {
                catPlotData.All.Reverse();
                CreateExtendedRegulonCategoryDataSheet(catPlotData, chartName);
            }

            // select the to re-activate the addin..
            this.RibbonUI.ActivateTab("TabGINtool");

            RemoveTask(TASKS.CATEGORY_CHART);

        }


        /// <summary>
        /// The routine that outputs the two bubble charts and worksheets to visualize the importance of the category/regulon
        /// </summary>
        /// <param name="aOutput"></param>
        /// <param name="aSummary"></param>        
        /// <param name="cat_Elements"></param>        
        /// <param name="splitNP"></param>
        private void RankingPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements)
        {            
            AddTask(TASKS.REGULON_CHART);

            SysData.DataTable _fc_BSU = ReformatResults(aOutput);

            cat_Elements = GetUniqueElements(cat_Elements);

            // HashSet ensures unique list
            HashSet<string> lRegulons = new HashSet<string>();

            foreach (SysData.DataRow row in aSummary.Rows)
                lRegulons.Add(row.ItemArray[0].ToString());

            SysData.DataView dataView = _fc_BSU.AsDataView();
            element_fc catPlotData;
            if (Properties.Settings.Default.useCat)
            {
                catPlotData = CatElements2ElementsFC(dataView, cat_Elements);
            }
            else
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements);

            (List<element_rank> plotData, List<summaryInfo> _all, List<summaryInfo> _pos, List<summaryInfo> _neg) = CreateRankingPlotData(catPlotData);

            int chartNr = Properties.Settings.Default.useCat ? NextWorksheet("CatRankPlot_") : NextWorksheet("RegRankPlot_");
            string chartName = (Properties.Settings.Default.useCat ? "CatRankPlot_" : "RegRankPlot_") + chartNr.ToString();            
            string chartNameBest = chartName.Replace("Plot_", "PlotBest_");

            (_, List<summaryInfo> _best) = CreateRankingDataSheet(catPlotData, _all, _pos, _neg);

            PlotRoutines.CreateRankingPlot2(plotData, chartName);

            if (!(_best is null))
            {                
                List<element_rank> _bestRankData = BubblePlotData(_best);
                PlotRoutines.CreateRankingPlot2(_bestRankData, chartNameBest+"_1", bestPlot: true);
                PlotRoutines.CreateRankingPlot2(_bestRankData, chartNameBest+"_2", bestPlot: true, bestNew:true);
            }


            this.RibbonUI.ActivateTab("TabGINtool");

            RemoveTask(TASKS.REGULON_CHART);

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
                        __values = _work.Select(x => x.p_average).ToArray();
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
        /// Transform the summaryinfo data to the highly specific (i.e. significance ranges) bubbleplot format ready result.
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        private List<element_rank> BubblePlotData(List<summaryInfo> info)
        {
            List<element_rank> element_Ranks = new List<element_rank>();

            // MAD values
            List<double> e1_m = new List<double>(), e2_m = new List<double>(), e3_m = new List<double>(), e4_m = new List<double>(), e5_m = new List<double>();
            // FC values
            List<double> e1_fc = new List<double>(), e2_fc = new List<double>(), e3_fc = new List<double>(), e4_fc = new List<double>(), e5_fc = new List<double>();
            // Counts
            List<int> e1_n = new List<int>(), e2_n = new List<int>(), e3_n = new List<int>(), e4_n = new List<int>(), e5_n = new List<int>();
            // CATEGORY/REGULON NAMES
            List<string> e1_s = new List<string>(), e2_s = new List<string>(), e3_s = new List<string>(), e4_s = new List<string>(), e5_s = new List<string>();
            // BEST GENE PERCENTAGES
            List<double> e1_p = new List<double>(), e2_p = new List<double>(), e3_p = new List<double>(), e4_p = new List<double>(), e5_p = new List<double>();



            foreach (summaryInfo sInfo in info)
            {
                List<double> _workfc = null;
                List<double> _workm = null;
                List<int> _workn = null;
                List<string> _works = null;
                List<double> _workp = null;

                if (sInfo.p_average < 0.06125 && sInfo.genes[0] != "")
                {
                    _workfc = e1_fc;
                    _workm = e1_m;
                    _workn = e1_n;
                    _works = e1_s;
                    _workp = e1_p;
                }

                if (sInfo.p_average >= 0.06125 && sInfo.p_average < 0.125 && sInfo.genes[0] != "")
                {
                    _workfc = e2_fc;
                    _workm = e2_m;
                    _workn = e2_n;
                    _works = e2_s;
                    _workp = e2_p;
                }


                if (sInfo.p_average >= 0.125 && sInfo.p_average < 0.25 && sInfo.genes[0] != "")
                {
                    _workfc = e3_fc;
                    _workm = e3_m;
                    _workn = e3_n;
                    _works = e3_s;
                    _workp = e3_p;
                }
                if (sInfo.p_average >= 0.25 && sInfo.p_average < 0.5 && sInfo.genes[0] != "")
                {
                    _workfc = e4_fc;
                    _workm = e4_m;
                    _workn = e4_n;
                    _works = e4_s;
                    _workp = e4_p;
                }


                if (sInfo.p_average >= 0.5 && sInfo.p_average <= 1 && sInfo.genes[0] != "")
                {
                    _workfc = e5_fc;
                    _workm = e5_m;
                    _workn = e5_n;
                    _works = e5_s;
                    _workp = e5_p;
                }

                if (_workfc != null)
                {

                    _workfc.Add(sInfo.fc_average);
                    _workm.Add(sInfo.fc_mad);
                    _workn.Add(sInfo.p_values != null ? sInfo.p_values.Length : 0);
                    _works.Add(StripText(sInfo.catName));
                    _workp.Add(sInfo.best_gene_percentage);
                }

            }


            element_rank e1 = new element_rank()
            {
                catName = "p<0.0625",
                average_fc = e1_fc.ToArray(),
                mad_fc = e1_m.ToArray(),
                nr_genes = e1_n.ToArray(),
                genes = e1_s.ToArray(),
                best_genes_percentage = e1_p.ToArray()
            };

            element_rank e2 = new element_rank()
            {
                catName = "0.0625>=p<0.125",
                average_fc = e2_fc.ToArray(),
                mad_fc = e2_m.ToArray(),
                nr_genes = e2_n.ToArray(),
                genes = e2_s.ToArray(),
                best_genes_percentage = e2_p.ToArray()
            };

            element_rank e3 = new element_rank()
            {
                catName = "0.125>=p<0.25",
                average_fc = e3_fc.ToArray(),
                mad_fc = e3_m.ToArray(),
                nr_genes = e3_n.ToArray(),
                genes = e3_s.ToArray(),
                best_genes_percentage = e3_p.ToArray()
            };

            element_rank e4 = new element_rank()
            {
                catName = "0.25>=p<0.5",
                average_fc = e4_fc.ToArray(),
                mad_fc = e4_m.ToArray(),
                nr_genes = e4_n.ToArray(),
                genes = e4_s.ToArray(),
                best_genes_percentage = e4_p.ToArray()
            };


            element_rank e5 = new element_rank()
            {
                catName = "0.5>=p=<1",
                average_fc = e5_fc.ToArray(),
                mad_fc = e5_m.ToArray(),
                nr_genes = e5_n.ToArray(),
                genes = e5_s.ToArray(),
                best_genes_percentage = e5_p.ToArray()

            };

            element_Ranks.Add(e1);
            element_Ranks.Add(e2);
            element_Ranks.Add(e3);
            element_Ranks.Add(e4);
            element_Ranks.Add(e5);


            return element_Ranks;
        }

        /// <summary>
        /// Order the ranking results by name and create bubble plot data
        /// </summary>
        /// <param name="theElements"></param>
        /// <returns></returns>
        private (List<element_rank>, List<summaryInfo>, List<summaryInfo>, List<summaryInfo>) CreateRankingPlotData(element_fc theElements)
        {

            List<summaryInfo> all_elements = SortedElements(theElements.All, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> pos_elements = SortedElements(theElements.Activated, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> neg_elements = SortedElements(theElements.Repressed, mode: SORTMODE.CATNAME, descending: false);

            return (BubblePlotData(theElements.All), all_elements, pos_elements, neg_elements);


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
        /// Create the tables for the observed 'best' ranking mode of operation of the regulon
        /// </summary>
        /// <param name="element_info"></param>
        /// <returns></returns>
        private (DataTable, List<summaryInfo>) BestElementScore(element_fc element_info)
        {

            // get best results..

            List<summaryInfo> _tmp = SortedElements(element_info.All, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> _output = new List<summaryInfo>();

            SysData.DataTable lTable = new SysData.DataTable("Elements");

            SysData.DataColumn regColumn = new SysData.DataColumn("Name", Type.GetType("System.String"));
            SysData.DataColumn statColumn1 = new SysData.DataColumn("Mode", Type.GetType("System.String"));
            SysData.DataColumn cntColumnA = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn percColumnA = new SysData.DataColumn("Percentage", Type.GetType("System.Int16"));
            SysData.DataColumn avgFCColumn1 = new SysData.DataColumn("AverageABSFC", Type.GetType("System.Double"));
            SysData.DataColumn madFCColumn1 = new SysData.DataColumn("MadABSFC", Type.GetType("System.Double"));
            SysData.DataColumn avgPColumn1 = new SysData.DataColumn("AverageP", Type.GetType("System.Double"));
            

            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(statColumn1);
            lTable.Columns.Add(cntColumnA);
            lTable.Columns.Add(percColumnA);
            lTable.Columns.Add(avgFCColumn1);
            lTable.Columns.Add(madFCColumn1);
            lTable.Columns.Add(avgPColumn1);


            for (int i = 0; i < _tmp.Count; i++)
            {
                bool swapped = false;
                SysData.DataRow lRow = lTable.Rows.Add();
                string catName = _tmp[i].catName;
                int totnrgenes = _tmp[i].genes.Length;
                summaryInfo _pos = element_info.Activated.GetCatValues(catName);
                summaryInfo _neg = element_info.Repressed.GetCatValues(catName);
                summaryInfo _si1 = _pos;
                summaryInfo _si2 = _neg;

                lRow["Name"] = StripText(catName);

                if (_pos.genes.Length < _neg.genes.Length)
                {
                    _si1 = _neg;
                    _si2 = _pos;
                    swapped = true;

                }

                int n1 = _si1.genes.Length;
                int n2 = _si2.genes.Length;

                if (n1 == n2) // check for highest FC
                {
                    if (Math.Abs(_si2.fc_average) > Math.Abs(_si1.fc_average))
                    {
                        _si1 = _neg;
                        swapped = !swapped;
                    }
                }                                

                lRow["Mode"] = swapped ? "repressed" : "activated";
                lRow["Count"] = _si1.genes.Length;
                if (totnrgenes > 0)
                    lRow["Percentage"] = Math.Round((double)_si1.genes.Length / (double)totnrgenes * 100);

                _si1.best_gene_percentage = Math.Round((double)_si1.genes.Length / (double)totnrgenes * 100);

                _output.Add(_si1);

                if (n1 > 0)
                {
                    lRow["AverageABSFC"] = _si1.fc_average;
                    lRow["MadABSFC"] = _si1.fc_mad;
                    lRow["AverageP"] = _si1.p_average;
                }

            }

            DataView _dv = lTable.DefaultView;
            _dv.Sort = "Name asc";

            return (_dv.ToTable(), _output);
        }

        /// <summary>
        /// Transform the ranking info to a summarized table 
        /// </summary>
        /// <param name="elements"></param>
        /// <param name="dirMode"></param>
        /// <returns></returns>
        private DataTable ElementsToTable(List<summaryInfo> elements, bool dirMode = false)
        {

            SysData.DataTable lTable = new SysData.DataTable("Elements");

            SysData.DataColumn regColumn = new SysData.DataColumn("Name", Type.GetType("System.String"));
            SysData.DataColumn dirColumn = new SysData.DataColumn("Direction", Type.GetType("System.String"));
            SysData.DataColumn cntColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
            SysData.DataColumn madColumn = new SysData.DataColumn("Mad", Type.GetType("System.Double"));
            SysData.DataColumn avgPColumn = new SysData.DataColumn("P_Average", Type.GetType("System.Double"));

            lTable.Columns.Add(regColumn);
            if (dirMode)
                lTable.Columns.Add(dirColumn);
            lTable.Columns.Add(cntColumn);
            lTable.Columns.Add(avgColumn);
            lTable.Columns.Add(madColumn);

            lTable.Columns.Add(avgPColumn);

            for (int r = 0; r < elements.Count; r++)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                string name = elements[r].catName;
                string[] names = name.Split('(');
                int hit = names[0].ToUpper().IndexOf("REGULON");
                string newname = hit == -1 ? names[0] : names[0].Substring(0, hit);

                lRow["Name"] = newname;
                if (dirMode)
                {
                    lRow["Direction"] = elements[r].fc_average > 0 ? "activation" : "repression";
                }
                lRow["Count"] = elements[r].p_values == null ? 0 : elements[r].p_values.Count();
                if (!(elements[r].fc_average is Double.NaN))
                    lRow["Average"] = elements[r].fc_average;
                if (!(elements[r].fc_mad is Double.NaN))
                    lRow["Mad"] = elements[r].fc_mad.ToString();
                if (!(elements[r].p_average is Double.NaN))
                    lRow["P_Average"] = elements[r].p_average;

            }

            return lTable;

        }

        /// <summary>
        /// Create a worksheet that contains the extended (=all genes) data per category/regulon. Used in spreading plot
        /// </summary>
        /// <param name="theElements"></param>
        /// <param name="chartName"></param>
        private void CreateExtendedRegulonCategoryDataSheet(element_fc theElements, string chartName)
        {

            string sheetName = chartName.Replace("Plot", "Tab");
            //aSheet.Name = sheetName;

            string catRegLabel = Properties.Settings.Default.useCat ? "Categgory" : "Regulon";

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            lNewSheet.Name = sheetName;
            //renameWorksheet(lNewSheet, catRegLabel+"_");

            DataTable lTable = ElementsToExtendedTable(theElements.All);

            lNewSheet.Cells[1, 1] = catRegLabel;
            lNewSheet.Cells[1, 2] = "Gene";
            lNewSheet.Cells[1, 3] = "FC";
            lNewSheet.Cells[1, 4] = "p-value";


            // starting from row 2


            FastDtToExcel(lTable, lNewSheet, 2, 1, lTable.Rows.Count + 1, lTable.Columns.Count);

        }

        /// <summary>
        /// Create the worksheet with data associated with the ranking bubble plots
        /// </summary>
        /// <param name="theElements"></param>
        /// <param name="all"></param>
        /// <param name="posSort"></param>
        /// <param name="negSort"></param>
        /// <returns></returns>

        private (Excel.Worksheet, List<summaryInfo>) CreateRankingDataSheet(element_fc theElements, List<summaryInfo> all, List<summaryInfo> posSort, List<summaryInfo> negSort)
        {
            string catRegLabel = Properties.Settings.Default.useCat ? "CatRankTab_" : "RegRankTab_";
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            RenameWorksheet(lNewSheet, catRegLabel);

            DataTable lTable = ElementsToTable(all);

            string catRegHeader = Properties.Settings.Default.useCat ? "Category" : "Regulon";
            string firstBlockHeader = Properties.Settings.Default.useCat ? "PLOT DATA" : "Without regulatory directionality";
            string secondBlockHeader = Properties.Settings.Default.useCat ? "POSITIVE FC" : "When regulator is activated";
            string thirdBlockHeader = Properties.Settings.Default.useCat ? "NEGATIVE FC" : "When regulator is repressed";
            string fourthBlockHeader = Properties.Settings.Default.useCat ? "COMBINED RESULTS" : "Best score";
            string FCheader = Properties.Settings.Default.useCat ? "Average FC" : "Average ABS(FC)";
            string MADheader = Properties.Settings.Default.useCat ? "MAD FC" : "MAD ABS(FC)";

            int hdrRow = 2;


            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[1, 5];
            Excel.Range rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
            rall.Merge();
            rall.Value = firstBlockHeader;
            rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            lNewSheet.Cells[hdrRow, 1] = catRegHeader;
            lNewSheet.Cells[hdrRow, 2] = "Nr Genes";
            lNewSheet.Cells[hdrRow, 3] = "Average FC";
            lNewSheet.Cells[hdrRow, 4] = "MAD FC";
            lNewSheet.Cells[hdrRow, 5] = "Average P";

            // Sort the data with ascending p-values
            DataView lView = lTable.DefaultView;

            // starting from row 2
            FastDtToExcel(lView.ToTable(), lNewSheet, hdrRow + 1, 1, lTable.Rows.Count + hdrRow, lTable.Columns.Count);

            if (!Properties.Settings.Default.useCat)
            {

                top = lNewSheet.Cells[1, 1];
                bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 5];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
                rall.Interior.TintAndShade = 0.8;
                rall.Interior.PatternTintAndShade = 0;



                top = lNewSheet.Cells[1, 7];
                bottom = lNewSheet.Cells[1, 11];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Merge();
                rall.Value = secondBlockHeader;
                rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

                lTable = ElementsToTable(posSort);

                lNewSheet.Cells[hdrRow, 7] = catRegHeader;
                lNewSheet.Cells[hdrRow, 8] = "Nr Genes";
                lNewSheet.Cells[hdrRow, 9] = FCheader;
                lNewSheet.Cells[hdrRow, 10] = MADheader;
                lNewSheet.Cells[hdrRow, 11] = "Average P";

                lView = lTable.DefaultView;

                FastDtToExcel(lView.ToTable(), lNewSheet, hdrRow + 1, 7, lTable.Rows.Count + hdrRow, lTable.Columns.Count + 6);

                top = lNewSheet.Cells[1, 7];
                bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 11];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
                rall.Interior.TintAndShade = 0.8;
                rall.Interior.PatternTintAndShade = 0;


                lTable = ElementsToTable(negSort);


                top = lNewSheet.Cells[1, 13];
                bottom = lNewSheet.Cells[1, 17];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Merge();
                rall.Value = thirdBlockHeader;
                rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

                lNewSheet.Cells[hdrRow, 13] = catRegHeader;
                lNewSheet.Cells[hdrRow, 14] = "Nr Genes";
                lNewSheet.Cells[hdrRow, 15] = FCheader;
                lNewSheet.Cells[hdrRow, 16] = MADheader;
                lNewSheet.Cells[hdrRow, 17] = "Average P";

                lView = lTable.DefaultView;

                FastDtToExcel(lView.ToTable(), lNewSheet, hdrRow + 1, 13, lTable.Rows.Count + hdrRow, lTable.Columns.Count + 12);

                top = lNewSheet.Cells[1, 13];
                bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 17];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
                rall.Interior.TintAndShade = 0.8;
                rall.Interior.PatternTintAndShade = 0;

                top = lNewSheet.Cells[1, 19];
                bottom = lNewSheet.Cells[1, 25];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Merge();
                rall.Value = fourthBlockHeader;

                rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                lNewSheet.Cells[hdrRow, 19] = catRegHeader;
                lNewSheet.Cells[hdrRow, 20] = "directon";
                lNewSheet.Cells[hdrRow, 21] = "Nr of genes";
                lNewSheet.Cells[hdrRow, 22] = "Percentage";
                lNewSheet.Cells[hdrRow, 23] = "Average ABS(FC)";
                lNewSheet.Cells[hdrRow, 24] = "MAD ABS(FC)";
                lNewSheet.Cells[hdrRow, 25] = "Average P";

                // Combine positive and negative mode results to obtain a 'best' result

                List<summaryInfo> _best;
                (lTable, _best) = BestElementScore(theElements);

                FastDtToExcel(lTable, lNewSheet, hdrRow + 1, 19, lTable.Rows.Count + hdrRow, lTable.Columns.Count + 18);

                top = lNewSheet.Cells[1, 19];
                bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 25];
                rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
                rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
                rall.Interior.TintAndShade = 0.8;
                rall.Interior.PatternTintAndShade = 0;

                return (lNewSheet, _best);
            }

            return (lNewSheet, null);
      

        }

        /// <summary>
        /// Select a csv file for regulon input
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, RibbonControlEventArgs e)
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
                    LoadWorksheets();
                    btLoad.Enabled = true;

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.referenceFile);
                    gLastFolder = fInfo.DirectoryName;
                }
            }
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
                    Properties.Settings.Default.operonFile = openFileDialog.FileName;
                    btnOperonFile.Label = Properties.Settings.Default.operonFile;
                    LoadOperonSheet();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.operonFile);
                    gLastFolder = fInfo.DirectoryName;

                }
            }
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
        }

        /// <summary>
        /// The text in the editbox for minimum p-value has changed. Update p-value dependent calculations.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void EditMinPval_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (double.TryParse(editMinPval.Text, out double val))
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
            var taskpane = TaskPaneManager.GetTaskPane("A", "GIN tool manual", () => new GINtaskpane(), SetTaskPaneVisbile);
            taskpane.Visible = !taskpane.Visible;
        }

        /// <summary>
        /// Store the visibility status of the task pane
        /// </summary>
        /// <param name="visible"></param>
        public void SetTaskPaneVisbile(bool visible)
        {
            tglTaskPane.Checked = visible;
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

            gOperonOutput = false;
            cbOperon.Checked = false;
            cbOperon.Enabled = false;
            Properties.Settings.Default.tblOperon = false;
        }

        /// <summary>
        /// Register the selection of ordering by FC (instead of P-value).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_OrderFC_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.useSort = cbOrderFC.Checked;
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
                    LoadCatFile();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.categoryFile);
                    gLastFolder = fInfo.DirectoryName;

                }
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
            cbUseRegulons.Checked = !cbUseCategories.Checked;
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
        }

        /// <summary>
        /// Show or hide the panels defined as settings
        /// </summary>
        /// <param name="show"></param>
        private void ShowSettingPannels(bool show)
        {
            grpReference.Visible = show;
            grpMap.Visible = show;
            grpUpDown.Visible = show;
            grpFC.Visible = show;
            grpCutOff.Visible = show;
            grpDirection.Visible = show;
        }


        /// <summary>
        /// enumeration of tasks that are defined
        /// </summary>
        public enum TASKS : int
        {
            /// <value>Nothing to do</value>
            READY = 0,
            /// <value>Regulon data is being read from file</value>
            LOAD_REGULON_DATA,
            /// <value>Operon data is being read from file</value>
            LOAD_OPERON_DATA,
            /// <value>Category is being read from file</value>
            LOAD_CATEGORY_DATA,
            /// <value>Genes are being mapped to regulons</value>
            MAPPING_GENES_TO_REGULONS,
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
            REGULON_CHART
        };

        public string[] taks_strings = new string[] { "Ready", "Load regulon data", "Load operon data", "Load category data", "Mapping genes to regulons", "Read sheet data", "Read sheet categorized data", "Update mapping table", "Update summary table", "Update combined table", "Update operon table", "Color cells", "Create category chart", "Create regulon chart" };

        /// <summary>
        /// enumeration of binary flags that can be set/unset.
        /// </summary>
        public enum UPDATE_FLAGS : byte
        {
            TSummary = 0b_0000_0001,
            TCombined = 0b_0000_0010,
            TOperon = 0b_0000_0100,
            TMapped = 0b_0000_1000,
            PRegulon = 0b_0001_0000,
            PDist = 0b_0010_0000,
            PCat = 0b_0100_0000,
            POperon = 0b_1000_0000,

            ///<value>FC dependency of multiple tables</value>
            FC_dependent = TCombined | POperon | TSummary,
            ///<value>P-value dependency of multiple tables</value>
            P_dependent = TCombined | POperon | TSummary,

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
            Properties.Settings.Default.tblOperon = cbOperon.Checked;
        }

        /// <summary>
        /// Clears the settings when category file is unset.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_ClearCatFile_Click(object sender, RibbonControlEventArgs e)
        {
            gCatOutput = false;
            cbUseCategories.Checked = false;
            cbUseCategories.Enabled = false;
            Properties.Settings.Default.useCat = false;
            cbUseRegulons.Checked = true;
            cbUseRegulons.Enabled = false;

            Properties.Settings.Default.categoryFile = "";
            btnCatFile.Label = "No file selected";



        }

        /// <summary>
        /// Register choice for p-values (instead of FCs).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void CheckBox_UsePValues_Click(object sender, RibbonControlEventArgs e)
        {
            cbUseFoldChanges.Checked = !cbUsePValues.Checked;
            Properties.Settings.Default.use_pvalues = cbUsePValues.Checked;
        }

        /// <summary>
        /// Register choice to use FCs (instead of P-Values)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_UseFoldChanges_Click(object sender, RibbonControlEventArgs e)
        {
            cbUsePValues.Checked = !cbUseFoldChanges.Checked;
            Properties.Settings.Default.use_pvalues = cbUsePValues.Checked;
        }


        /// <summary>
        /// What to do when data is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Select_Click(object sender, RibbonControlEventArgs e)
        {

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

                (gOutput, gList) = GenerateOutput();

                if (gOutput is null || gList is null)
                    return;

                btApply.Enabled = true;
                btPlot.Enabled = true;
                EnableOutputOptions(true);

                UnSetFlags(UPDATE_FLAGS.TMapped);

            }

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
            Properties.Settings.Default.useCat = !cbUseRegulons.Checked;
            cbUseCategories.Checked = !cbUseRegulons.Checked;
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
