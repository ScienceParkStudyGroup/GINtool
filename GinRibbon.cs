#undef CLICK_CHART // check to include clickable chart and events.. only if object storage is an option.

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;


namespace GINtool
{
  
    public partial class GinRibbon
    {
        #region inits

        string gLastFolder = "";
        bool gOperonOutput = false;
        bool gCatOutput = false;
        //bool gpValueUpdate = true;
        //bool gRegulonPlot = true;

#if CLICK_CHART
        List<chart_info> gCharts = new List<chart_info>();
#endif

        byte gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;


        List<TASKS> gTasks = new List<TASKS>();

        // int[] updateFlags = new int[3] { 1, 1, 1 };

        //bool SummaryTableNeedsUpdating = true;
        //bool CombinedTableNeedsUpdating = true;
        //bool OperonTableNeedsUpdating = true;
        //bool MappedTableNeedsUpdating = true;
        //bool RegulonPlotsNeedsUpdating = true;
        //bool CategoryPlotNeedsUpdating = true;
        //bool DistributionPlotNeedsUpdating = true;
        //bool RegulonPlotNeedsUpdating = true;


        int maxGenesPerOperon = 1;

        SysData.DataTable gRefWB = null; // RegulonData .. rename later
        SysData.DataTable gRefStats = null;
        SysData.DataTable gRefOperons = null;
        SysData.DataTable gCategories = null;
        string[] gColNames = null;
        string gCategoryGeneColumn = "locus_tag"; // the fixed column name that refers to the genes inthe category csv file
        Excel.Application gApplication = null;

        //bool gGeneratePlots = false;
        //bool gQPlot = false;
        //bool gNeedsUpdating = true;
        //bool gOrderAscending = true;
        bool gSortResults = false;
        
        //bool gUseCatOutput = false;

        static List<string> gAvailItems = null;
        static List<string> gUpItems = null;
        static List<string> gDownItems = null;

        List<int> gExcelErrorValues = null;

        //string gInputRange = "";

        string gOldRangeBSU = "";
        string gOldRangeP = "";
        string gOldRangeFC = "";

        Excel.Range gRangeBSU;
        Excel.Range gRangeFC;
        Excel.Range gRangeP;
        List<FC_BSU> gOutput = null;
        SysData.DataTable gSummary = null;
        List<BsuRegulons> gList = null;
        SysData.DataTable gCombineInfo = null;

        //int gDDcatLevel = 1;
        
        
       // private bool gPlotDistribution = false;
     //   private bool gPlotClustered = false; // is category plot

        //PlotRoutines gEnrichmentAnalysis;

#endregion

#region database_utils
        private List<string> propertyItems(string property)
        {
            StringCollection myCol = (StringCollection)Properties.Settings.Default[property];

            if (myCol != null)
                return myCol.Cast<string>().ToList();

            return new List<string>();
        }


        private void storeValue(string property, List<string> aValue)
        {

            StringCollection collection = new StringCollection();
            collection.AddRange(aValue.ToArray());

            Properties.Settings.Default[property] = collection;
        }

        private SysData.DataTable GetDistinctRecords(SysData.DataTable dt, string[] Columns)
        {
            return dt.DefaultView.ToTable(true, Columns);
        }

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
#endregion

        private bool LoadCategoryData()
        {
            if (Properties.Settings.Default.categoryFile.Length == 0 || Properties.Settings.Default.catSheet.Length == 0)
                return false;

            //gApplication.StatusBar = "Load category data";
            //gApplication.EnableEvents = false;

            AddTask(TASKS.LOAD_CATEGORY_DATA);


            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.catSheet, Properties.Settings.Default.categoryFile, 1, 1);
            gCategories = new SysData.DataTable("Categories");
            gCategories.CaseSensitive = false;
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



            string[] lcols = new string [] {"cat1_int","cat2_int","cat3_int","cat4_int","cat5_int"};
            string[] ulcols = new string[] { "ucat1_int", "ucat2_int", "ucat3_int", "ucat4_int", "ucat5_int" };

            foreach (SysData.DataRow lRow in _tmp.Rows)
            {
                object[] lItems = lRow.ItemArray;
                SysData.DataRow lNewRow = gCategories.Rows.Add();
                for (int i = 0; i < lItems.Length; i++)
                {
                    lNewRow["catid"] = lItems[0];
                    string[] splits = lItems[0].ToString().Split(' ');
                    lNewRow["catid_short"] = splits[splits.Count()-1];
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
                    for (int j = llItems.Length; j < 5; j++)
                    {
                        //lNewRow[lcols[j]] = 0;
                        //lNewRow[ulcols[j]] = 0;
                    }

                    int offset = 0;
                    for (int j = 0; j < llItems.Length; j++)
                    {                       
                        lNewRow[ulcols[j]] = offset + ((Int32)lNewRow[lcols[j]]) * Math.Pow(10, 5 - j);
                        offset = (Int32)lNewRow[ulcols[j]];
                    }

                }
            }
            //gApplication.EnableEvents = true;
            //gApplication.StatusBar = "Ready";

            RemoveTask(TASKS.LOAD_CATEGORY_DATA);

            return gCategories.Rows.Count > 0;
        }

        private bool LoadOperonData()
        {
          
            if (Properties.Settings.Default.operonFile.Length == 0 || Properties.Settings.Default.operonSheet.Length == 0)            
                return false;

            AddTask(TASKS.LOAD_OPERON_DATA);
            //gApplication.EnableEvents = false;
            //gApplication.StatusBar = "Load operon data";

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.operonSheet, Properties.Settings.Default.operonFile, 1, 1);
            gRefOperons = new SysData.DataTable("OPERONS");
            gRefOperons.CaseSensitive = false;
            gRefOperons.Columns.Add("operon", Type.GetType("System.String"));
            gRefOperons.Columns.Add("gene", Type.GetType("System.String"));
            gRefOperons.Columns.Add("op_id", Type.GetType("System.Int32"));

            int _op_id = 0;

            foreach(SysData.DataRow lRow in _tmp.Rows)
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
            //gApplication.EnableEvents = true;
            //gApplication.StatusBar = "Ready";

            RemoveTask(TASKS.LOAD_OPERON_DATA);
            return gRefOperons.Rows.Count>0;
        }

        private bool LoadData()
        {
            //gApplication.StatusBar = "Load regulon/gene mappings";
            //gApplication.EnableEvents = false;

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
            //gApplication.EnableEvents = true;
            //gApplication.StatusBar = "Ready";

            RemoveTask(TASKS.LOAD_REGULON_DATA);
            return gRefWB != null ? true : false;
        }

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


        private void InitFields(bool enable=false)
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


        private void LoadButtonStatus()
        {
            gApplication = Globals.ThisAddIn.GetExcelApplication();
            btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;

            if (btnRegulonFileName.Label.Length>0)
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
            cbUsePValues.Checked = Properties.Settings.Default.use_pvalues;
            cbUseFoldChanges.Checked = !Properties.Settings.Default.use_pvalues;


          
        }

        private void GinRibbon_Load(object sender, RibbonUIEventArgs e)
        {

            LoadButtonStatus();

            if (Properties.Settings.Default.operonFile.Length == 0)
                btnOperonFile.Label = "No file selected";

            gAvailItems = propertyItems("directionMapUnassigned");
            gUpItems = propertyItems("directionMapUp");
            gDownItems = propertyItems("directionMapDown");

            InitFields();
           
            PlotRoutines.theApp = gApplication;

            EnableOutputOptions(false);

            gExcelErrorValues = ((int[])Enum.GetValues(typeof(ExcelUtils.CVErrEnum))).ToList();
          
            btLoad.Enabled = System.IO.File.Exists(Properties.Settings.Default.referenceFile);

        }

        private Excel.Range GetActiveCell()
        {
            if (gApplication != null)
            {
                try { return (Excel.Range)gApplication.Selection; }
                catch (Exception e) { return null; }
                
            }
            return null;
        }

        private void EnableOutputOptions(bool enable)
        {
            ebLow.Enabled = enable;
            ebMid.Enabled = enable;
            ebHigh.Enabled = enable;
            editMinPval.Enabled = enable;            
            //splitButton3.Enabled = enable;

            //cbUseCategories.Enabled = enable;
            cbMapping.Enabled = enable;
            cbSummary.Enabled = enable;
            cbCombined.Enabled = enable;

            cbClustered.Enabled = enable;
            cbDistribution.Enabled = enable;
            chkRegulon.Enabled= enable;

            cbOrderFC.Enabled = enable;
            cbUseCategories.Enabled = enable && gCatOutput;
            cbUseRegulons.Enabled = enable;

            cbOperon.Enabled = enable && gOperonOutput;

            cbUsePValues.Enabled = enable;
            cbUseFoldChanges.Enabled = enable;

            cbAscending.Enabled = enable;
            cbDescending.Enabled = enable;

            //splitbtnEA.Enabled = enable;
        }


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
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return null;
        }

        private void ConditionFormatRange(Excel.Range columnRange)
        {

            Excel.FormatConditions fcs = columnRange.FormatConditions;

            var formatCondition = fcs.Add(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlDatabar);

            formatCondition.MinPoint.Modify(Microsoft.Office.Interop.Excel.XlConditionValueTypes.xlConditionValueAutomaticMin);
            formatCondition.MaxPoint.Modify(Microsoft.Office.Interop.Excel.XlConditionValueTypes.xlConditionValueAutomaticMax);


            formatCondition.BarFillType = Microsoft.Office.Interop.Excel.XlGradientFillType.xlGradientFillPath;
            formatCondition.Direction = Microsoft.Office.Interop.Excel.Constants.xlContext;
            formatCondition.NegativeBarFormat.ColorType = Microsoft.Office.Interop.Excel.XlDataBarNegativeColorType.xlDataBarColor;

            formatCondition.BarColor.Color = 8700771; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            formatCondition.BarColor.TintAndShade = 0; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);


            formatCondition.BarBorder.Color.Color = 8700771; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            formatCondition.BarBorder.Type = Microsoft.Office.Interop.Excel.XlDataBarBorderType.xlDataBarBorderSolid;

            formatCondition.NegativeBarFormat.BorderColorType = Microsoft.Office.Interop.Excel.XlDataBarNegativeColorType.xlDataBarColor;
            formatCondition.NegativeBarFormat.Parent.BarBorder.Type = Microsoft.Office.Interop.Excel.XlDataBarBorderType.xlDataBarBorderSolid;

            formatCondition.AxisPosition = Microsoft.Office.Interop.Excel.XlDataBarAxisPosition.xlDataBarAxisAutomatic;

            formatCondition.AxisColor.Color = 0; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            formatCondition.AxisColor.TintAndShade = 0;

            formatCondition.NegativeBarFormat.Color.Color = 255; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
            formatCondition.NegativeBarFormat.Color.TintAndShade = 0;

            formatCondition.NegativeBarFormat.BorderColor.Color = 255; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
            formatCondition.NegativeBarFormat.BorderColor.TintAndShade = 0;
        }

        private bool iserrorCell(object obj)
        {
            
            return (obj is Int32) && gExcelErrorValues.Contains((Int32)obj);
        }

        private List<BsuRegulons> QueryResultTable(List<Excel.Range> theCells)
        {

            AddTask(TASKS.MAPPING_GENES_TO_REGULONS);

            object[,] rangeBSU = theCells[2].Value2;
            object[,] rangeFC = theCells[1].Value2;
            object[,] rangeP = theCells[0].Value2;


            List<BsuRegulons> lList = new List<BsuRegulons>();
            
            for(int _r=1;_r<=rangeBSU.Length;_r++)
            {
                string lBSU;
                double lFC = 0;
                double lPvalue = 1;
                BsuRegulons lMap = null;


                //if (c.Columns.Count == 3)
                {
                    
                    lBSU = rangeBSU[_r,1].ToString();

                    if (!iserrorCell(rangeP[_r,1]))
                        if (!Double.TryParse(rangeP[_r,1].ToString(), out lPvalue))
                            lPvalue = 1;

                    if (!iserrorCell(rangeFC[_r,1]))
                        if (!Double.TryParse(rangeFC[_r,1].ToString(), out lFC))
                            lFC = 0;

                    //lBSU = value[1, 3].ToString();
                    lMap = new BsuRegulons(lFC, lPvalue, lBSU);
                }

                if (lMap.BSU.Length > 0)
                {
                    SysData.DataRow[] results = Lookup(lMap.BSU);

                    if (results.Length > 0)
                    {
                        string gene = results[0][Properties.Settings.Default.referenceGene].ToString();
                        lMap.GENE = gene;

                        for (int r = 0; r < results.Length; r++)
                        {
                            string item = results[r][Properties.Settings.Default.referenceRegulon].ToString();
                            string direction = results[r][Properties.Settings.Default.referenceDIR].ToString();


                            if (item.Length > 0) // loop over found regulons
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
   

        //private List<BsuRegulons> QueryResultTable(List<Excel.Range> theCells)
        //{
            
        //    AddTask(TASKS.MAPPING_GENES_TO_REGULONS);

        //    List<BsuRegulons> lList = new List<BsuRegulons>();

        //    foreach (Excel.Range c in theCells[0].Rows)
        //    {
        //        string lBSU;                
        //        double lFC = 0;
        //        double lPvalue = 1;
        //        BsuRegulons lMap = null;
               
        //        if (c.Columns.Count == 3)
        //        {
        //            object[,] value = c.Value2;
                    
        //            // first check if the cell contains an erroneous value, if not then try to parse the value or reset to default

        //            if (!iserrorCell(value[1, 1]))
        //                if (!Double.TryParse(value[1, 1].ToString(), out lPvalue))
        //                    lPvalue = 1;

        //            if (!iserrorCell(value[1, 2]))
        //                if (!Double.TryParse(value[1, 2].ToString(), out lFC))
        //                    lFC = 0;
                                        
        //            lBSU = value[1, 3].ToString();
        //            lMap = new BsuRegulons(lFC, lPvalue, lBSU);                    
        //        }
                
        //        if (lMap.BSU.Length > 0)
        //        {
        //            SysData.DataRow[] results = Lookup(lMap.BSU);

        //            if (results.Length > 0)
        //            {
        //                string gene = results[0][Properties.Settings.Default.referenceGene].ToString();
        //                lMap.GENE = gene;

        //                for (int r = 0; r < results.Length; r++)
        //                {
        //                    string item = results[r][Properties.Settings.Default.referenceRegulon].ToString();
        //                    string direction = results[r][Properties.Settings.Default.referenceDIR].ToString();
                            

        //                    if (item.Length > 0) // loop over found regulons
        //                    {
        //                        lMap.REGULONS.Add(item);
                                
        //                        if (gUpItems.Contains(direction))                                
        //                            lMap.UP.Add(r);                                    
                                
        //                        if (gDownItems.Contains(direction))                                
        //                            lMap.DOWN.Add(r);                                                                    
        //                    }
        //                }
        //            }
        //        }

        //        lList.Add(lMap);
        //    }

        //    RemoveTask(TASKS.MAPPING_GENES_TO_REGULONS);
        //    return lList;

        //}
     
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



        public string RangeAddress(Excel.Range rng)
        {
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1,Type.Missing, Type.Missing);
        }
        public string CellAddress(Excel.Worksheet sht, int row, int col)
        {
            return RangeAddress(sht.Cells[row, col]);
        }


        private string GetStatusTask(TASKS task)
        {
            return taks_strings[(int)task];
        }

        

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

        private void AddTask(TASKS newTask)
        {
            gTasks.Add(newTask);
            SetStatus(newTask);
        }

        private void RemoveTask(TASKS taskReady)
        {
            gTasks.Remove(taskReady);
            if (gTasks.Count == 0 || gTasks[0] == TASKS.READY)
                SetStatus(TASKS.READY);
            else
                SetStatus(gTasks.Last());
        }


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

        // the main routine after mouse selection update // generates mapping output.. should be de-coupled (update data & mapping output)
        private (List<FC_BSU>, List<BsuRegulons>) GenerateOutput(bool suppressOutput=false)
        {

            AddTask(TASKS.READ_SHEET_DATA);

            //Excel.Range theInputCells = GetActiveCell();
            List<Excel.Range> theInputCells = new List<Excel.Range>();
            
            theInputCells.Add(gRangeP);
            theInputCells.Add(gRangeFC);
            theInputCells.Add(gRangeBSU);
            

            Excel.Worksheet theSheet = GetActiveSheet();

         


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
            //int startC = theInputCells.Column;
            //int startR = theInputCells.Row;

            // from now always assume 3 columns.. p-value, fc, bsu

            // generate the results for outputting the data and summary
            try
            {
                List<BsuRegulons> lResults = QueryResultTable(theInputCells);
               
            
                List<FC_BSU> lOutput = new List<FC_BSU>();

                for (int r = 0; r < nrRows; r++)
                    for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                    {
                        int val = 0;
                        if (lResults[r].UP.Contains(c))
                            val = 1;
                        if (lResults[r].DOWN.Contains(c))
                            val = -1;

                        lOutput.Add(new FC_BSU(lResults[r].FC, lResults[r].REGULONS[c], val, lResults[r].PVALUE,lResults[r].GENE));
                    }

                RemoveTask(TASKS.READ_SHEET_DATA);

                return (lOutput, lResults);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Are you shure the columns do not contain text?");
                RemoveTask(TASKS.READ_SHEET_DATA);

                return (null, null);
            }        
           
        }
       

        private void CreateMappingSheet(List<BsuRegulons> bsuRegulons)
        {
            var lOut = PrepareResultTable(bsuRegulons);

            SysData.DataTable lTable = lOut.Item1;
            SysData.DataTable clrTbl;

            AddTask(TASKS.UPDATE_MAPPED_TABLE);
            //gApplication.StatusBar = "Creating mapping sheet";

            int nrRows = lTable.Rows.Count;
            //int startC = 1;
            int startR = 2;

            // from now always assume 3 columns.. p-value, fc, bsu
            int offsetColumn = 1;


            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "Mapped_");

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

            FastDtToExcel(lTable, lNewSheet, startR, offsetColumn, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);

            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[lTable.Rows.Count + 1, lTable.Columns.Count];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();
           //all.Select().Rows.AutoFit();

            clrTbl = lOut.Item2;
            ColorCells(clrTbl, lNewSheet, startR, offsetColumn + 5, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);
            
            RemoveTask(TASKS.UPDATE_MAPPED_TABLE);
            
        }

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


        private void CreateSummarySheet(SysData.DataTable theTable)
        {

            AddTask(TASKS.UPDATE_SUMMARY_TABLE);

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "Summary_");

            int col = 1;
            
            
            Excel.Range top = lNewSheet.Cells[1, 4];
            Excel.Range bottom = lNewSheet.Cells[1, 11];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Observed Counts and directions";
            all.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            top = lNewSheet.Cells[1, 13];
            bottom = lNewSheet.Cells[1, 20];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Percentage";
            all.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            top = lNewSheet.Cells[1, 21];
            bottom = lNewSheet.Cells[1, 22];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Logical direction";
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
            int colGreen = col;

            lNewSheet.Cells[2, col++] = string.Format("UP >{0}", Properties.Settings.Default.fcHIGH);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >{1}", Properties.Settings.Default.fcHIGH, Properties.Settings.Default.fcMID);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >{1}", Properties.Settings.Default.fcMID, Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("UP <={0} & >0", Properties.Settings.Default.fcLOW);
            
            lNewSheet.Cells[2, col++] = string.Format("DOWN <0 & >=-{0}", Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <-{0} & >=-{1}", Properties.Settings.Default.fcMID, Properties.Settings.Default.fcLOW);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <=-{0} & >=-{1}", Properties.Settings.Default.fcHIGH, Properties.Settings.Default.fcMID);
            lNewSheet.Cells[2, col++] = string.Format("DOWN <-{0}", Properties.Settings.Default.fcHIGH);

            lNewSheet.Cells[2, col++] = "DOWN";
            lNewSheet.Cells[2, col++] = "UP";
           
            // starting from row 3


            FastDtToExcel(theTable, lNewSheet, 3, 1, theTable.Rows.Count + 2, theTable.Columns.Count);


            // color cells here

            
            top = lNewSheet.Cells[3, colGreen];
            bottom = lNewSheet.Cells[theTable.Rows.Count + 2, colGreen+4];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);

            top = lNewSheet.Cells[3, colGreen+4];
            bottom = lNewSheet.Cells[theTable.Rows.Count + 2, colGreen + 4+3];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);

            // set number format

            top = lNewSheet.Cells[3, 13];
            bottom = lNewSheet.Cells[2 + theTable.Rows.Count, 22];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.NumberFormat = "###%";


            // fit the width of the columns
            top = lNewSheet.Cells[1, 1];
            bottom = lNewSheet.Cells[theTable.Rows.Count+2, theTable.Columns.Count];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();


            RemoveTask(TASKS.UPDATE_SUMMARY_TABLE);

        }

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


        private (int,int, int) CalculateFPRatio(SysData.DataRow[] aRow)
        {
            int nrUP=0, nrDOWN = 0, nrTot=0;

            // aRow from an FC_BSU table
            for (int i = 0; i < aRow.Length; i++)
            {                
                double fcGene = (double)aRow[i]["FC"];
                int dirBSU = (int)aRow[i]["DIR"];
                double lowValue = Properties.Settings.Default.fcLOW;

                // if upregulated
                if (dirBSU < 0)
                {
                    if (fcGene < -lowValue)
                    {
                        nrUP += 1;
                        nrTot += 1;
                    }
                    
                }
                if (dirBSU > 0)
                {
                    if (fcGene > lowValue)
                    {
                        nrUP += 1;
                        nrTot += 1;
                    }
                }
                // if downregulated
                if (dirBSU > 0)
                {
                    if (fcGene < -lowValue)
                    {
                        nrDOWN += 1;
                        nrTot += 1;
                    }

                }
                if (dirBSU < 0)
                {
                    if (fcGene > lowValue)
                    {
                        nrDOWN += 1;
                        nrTot += 1;
                    }
                }

            }

            return (nrUP, nrDOWN, nrTot);
        }


        private List<string> listSheets()
        {
            // get a list of all sheet names
            List<string> _sheets = new List<string>();

            foreach (var sheet in gApplication.Sheets)
            {
                if (sheet is Excel.Chart)                
                    _sheets.Add(((Excel.Chart)sheet).Name);                
                else
                    _sheets.Add(((Excel.Worksheet)sheet).Name);
            }

            return _sheets;
                
        }


        private int nextWorksheet(string wsBase)
        {
            // create a sheetname starting with wsBase
            List<string> currentSheets = listSheets();

            string sheetName = wsBase.Replace("Plot", "Tab");
            string chartName = wsBase.Replace("Tab", "Plot");

            int s = 1;
            while (currentSheets.Contains(string.Format("{0}{1}", chartName, s)) || currentSheets.Contains(string.Format("{0}{1}", sheetName, s)))  
                s += 1;

           return s;            
        }

        private int renameWorksheet(object aSheet, string wsBase)
        {
            // create a sheetname starting with wsBase
            List<string> currentSheets = listSheets();
            int s = 1;
            while (currentSheets.Contains(string.Format("{0}_{1}", wsBase, s)))
                s += 1;

            if (aSheet is Excel.Chart)
            {
                Excel.ChartObject chartObject = (Excel.ChartObject)((Excel.Chart)aSheet).Parent;
                //chartObject.
                //((Excel.Chart)aSheet).Name = string.Format("{0}_{1}", wsBase, s);
            }
            else
                ((Excel.Worksheet)(aSheet)).Name = string.Format("{0}_{1}", wsBase, s);

            return s;
        }        

        // from https://stackoverflow.com/questions/665754/inner-join-of-datatables-in-c-sharp
        private DataTable JoinDataTables(DataTable t1, DataTable t2, params Func<DataRow, DataRow, bool>[] joinOn)
        {
            DataTable result = new DataTable();
            foreach (DataColumn col in t1.Columns)
            {
                if (result.Columns[col.ColumnName] == null)
                    result.Columns.Add(col.ColumnName, col.DataType);
            }
            foreach (DataColumn col in t2.Columns)
            {
                if (result.Columns[col.ColumnName] == null)
                    result.Columns.Add(col.ColumnName, col.DataType);
            }
            foreach (DataRow row1 in t1.Rows)
            {
                var joinRows = t2.AsEnumerable().Where(row2 =>
                {
                    foreach (var parameter in joinOn)
                    {
                        if (!parameter(row1, row2)) return false;
                    }
                    return true;
                });
                foreach (DataRow fromRow in joinRows)
                {
                    DataRow insertRow = result.NewRow();
                    foreach (DataColumn col1 in t1.Columns)
                    {
                        insertRow[col1.ColumnName] = row1[col1.ColumnName];
                    }
                    foreach (DataColumn col2 in t2.Columns)
                    {
                        insertRow[col2.ColumnName] = fromRow[col2.ColumnName];
                    }
                    result.Rows.Add(insertRow);
                }
            }
            return result;
        }



        //private Excel.Worksheet CategoryData(SysData.DataTable aTable, List<cat_elements> cat_s)
        //{

        //    gApplication.StatusBar = "Generating categorized sheet for plotting data";
        //    gApplication.ScreenUpdating = false;
        //    gApplication.DisplayAlerts = false;
        //    gApplication.EnableEvents = false;

        //    Excel.Worksheet aSheet = gApplication.Worksheets.Add();


        //    //string[] cols = new string[] { string.Format("ucat{0}_int",gDDcatLevel.ToString()) };
        //    //SysData.DataTable dt_unique = GetDistinctRecords(gCategories, cols);
            

        //    //List<string> lCategories = new List<string>();
        //    //for (int i = 0; i < dt_unique.Rows.Count; i++)
          

        //    //DataTable newTable = GinRibbon.DtTbl(new DataTable[] {lTable,aTable});

        //    //DataTable newTable = lTable.Merge(aTable);

        //    int nrGenes = aTable.Rows.Count;
        //    //int nrRegulons = lCategories.Count;

        //    List<float[]> fc = new List<float[]>();

        //    SysData.DataView dataView = aTable.AsDataView();







        //    double MMAX = (double)(float)aTable.Rows[0]["FC"];
        //    double MMIN = (double)(float)aTable.Rows[0]["FC"];
        //    List<double> meanFC = new List<double>();


        //    int nrcat = 0;
        //    foreach (cat_elements cs in cat_s)
        //    {


        //        string categories = string.Join(",", cs.elements.ToArray());
        //        categories = string.Join(",", categories.Split(',').Select(x => $"'{x}'"));


        //        DataTable lTable = gCategories.Select(string.Format("catid_short in= {0}",categories)).CopyToDataTable();

        //        DataTable _dataTable = JoinDataTables(lTable, aTable,
        //           (row1, row2) =>
        //           row1.Field<string>(gCategoryGeneColumn) == row2.Field<string>("Gene"));

        //        if (_dataTable.Rows.Count == 0)
        //            continue;

        //        _dataTable = GetDistinctRecords(_dataTable,new string[] { "Gene" });

        //        //DataRow[] _lResult = gCategories.Select(string.Format("ucat{0}_int = {1}", gDDcatLevel.ToString(), row[0].ToString()));

        //        //lCategories.Add(string.Format("{0}", _lResult[0].ItemArray[gDDcatLevel+2]));
        //        //nrcat++;


        //        int nrRows = _dataTable.Rows.Count;
        //        float[] vs = new float[nrRows];
        //        int[] ys = new int[nrRows];
        //        for (int _r = 0; _r < nrRows; _r++)
        //        {
        //            double _val = (double)(float)_dataTable.Rows[_r]["FC"];
        //            if (_val > MMAX) { MMAX = _val; }
        //            if (_val < MMIN) { MMIN = _val; }
        //            vs[_r] = (float)_val;
        //            ys[_r] = fc.Count;

        //        }
        //        meanFC.Add(vs.Average());
        //        fc.Add(vs);
        //    }


        //    int nrRegulons = cat_s.Count();

        //    if (nrRegulons>255)
        //    {
        //        MessageBox.Show("No more than 255 series can be plotted, please select fewer categories");
        //        return null;
        //    }


        //    int[] sortedEntries = Enumerable.Range(0, nrRegulons).ToArray();

        //    if (gOrderAscending)
        //    {
        //        double[] __values = meanFC.ToArray();
        //        var sortedEntriesPairs = __values
        //            .Select((x, i) => new KeyValuePair<double, int>(x, i))
        //            .OrderBy(x => x.Key)
        //            .ToList();

        //        sortedEntries = sortedEntriesPairs.Select(x => x.Value).ToArray();
        //    }


        //    int totrow = 2;
        //    for (int i = 0; i < cat_s.Count; i++)
        //    {
        //        aSheet.Cells[1, 1] = "yoffset";
        //        aSheet.Cells[1, i + 2] = cat_s[i].elements.ToArray()[sortedEntries[i]];

        //        for (int j = 0; j < fc[sortedEntries[i]].Length; j++)
        //        {
        //            aSheet.Cells[totrow, 1] = i + 0.5;
        //            aSheet.Cells[totrow++, i + 2] = fc[sortedEntries[i]][j];
        //        }
        //    }


        //    gApplication.ScreenUpdating = true;
        //    gApplication.DisplayAlerts = true;
        //    gApplication.EnableEvents = true;

        //    gApplication.StatusBar = "Ready";

        //    return aSheet;

        //}

        private void CreateOperonSheet(SysData.DataTable table)
        {
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "Operon_");


            int maxNrGenes = Int32.Parse(table.Compute("max([nrgenes])",string.Empty).ToString());

            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

            int firstRow = 1;
            int firstCol = 1;
            //int lastCol = table.Columns.Count + firstCol;
            int lastCol = (maxNrGenes + 4) + firstCol;
            int lastRow = table.Rows.Count + firstRow;

            Excel.Range top = lNewSheet.Cells[firstRow, firstCol];
            Excel.Range bottom = lNewSheet.Cells[lastRow, lastCol];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            int col = 1;
            lNewSheet.Cells[1, col++] = "BSU";
            lNewSheet.Cells[1, col++] = "FC";
            lNewSheet.Cells[1, col++] = "P-Value";

            lNewSheet.Cells[1, col++] = "Gene";
            lNewSheet.Cells[1, col++] = "Operon Name";
            lNewSheet.Cells[1, col++] = "Nr operons";
            lNewSheet.Cells[1, col++] = "Nr genes";
            lNewSheet.Cells[1, col++] = "Operon";

            for (int c=0;c< maxNrGenes; c++)
            {
                string colHeader = string.Format("FC Gene #{0}", c + 1);
                lNewSheet.Cells[1, c+col] = colHeader;
            }

            FastDtToExcel(table, lNewSheet, 2, 1, table.Rows.Count + 1, maxNrGenes + 4);


            top = lNewSheet.Cells[1, 1];
            bottom = lNewSheet.Cells[table.Rows.Count + 1, maxNrGenes + 4];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();
            all.Rows.AutoFit();


            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;
        }

        private void CreateCombinedSheet(SysData.DataTable aTable)
        {
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "Combined_");

            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

            int firstRow = 1;
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
            
            //if(gOperonOutput)
            //    lNewSheet.Cells[1, col++] = "OPERON(S)";


            // determine the maximum number of regulons from the table that wass passed

            string lastColumn = aTable.Columns[aTable.Columns.Count - 1].ColumnName;
            lastColumn = lastColumn.Replace("Regulon_", "");
            int maxRegulons = Int16.Parse(lastColumn);

            for (int r = 0; r < aTable.Rows.Count; r++)
            {
                SysData.DataRow clrRow = aTable.Rows[r];
                for (int c = 0; c < clrRow.ItemArray.Length; c++)
                {
                    Excel.Range lR = all.Cells[r + 2, c + 1];
                    int UpPos = clrRow[c].ToString().IndexOf("#");
                    int DownPos = clrRow[c].ToString().IndexOf("@");
                    int UpColor = clrRow[c].ToString().IndexOf('&');
                    int DownColor = clrRow[c].ToString().IndexOf('!');

                    lR.Value = clrRow[c];

                    if (clrRow[c].ToString().Length == 0)
                        continue;
                    
                    if (UpPos == -1 && DownPos == -1)
                        continue;

                    if (UpPos > 0)
                    {
                        Excel.Characters lChar = lR.Characters[UpPos+1, 1];                        
                        lChar.Text = "á"; // the arrow up
                        lChar.Font.Name = "Wingdings";

                        if (UpColor > 0)
                        {
                            // delete the & symbol
                            lR.Characters[UpColor+1,1].Delete();
                            lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                        }
                        else
                        {
                            // delete the ! symbol
                            lR.Characters[DownColor+1,1].Delete();
                            lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                        }
                    }
                    else
                    {
                        Excel.Characters lChar = lR.Characters[DownPos + 1, 2];
                        lChar.Text = "â "; // the arrow down
                        lChar.Font.Name = "Wingdings";
                        if (DownColor > 0)
                        {
                            // delete the ! symbol
                            lR.Characters[DownColor+1,1].Delete();
                            lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                        }
                        else
                        {
                            // delete the & symbol
                            lR.Characters[UpColor+1,1].Delete();
                            lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                        }
                            
                    }                                            
                }
            }

            for (int c = 0; c < maxRegulons; c++)
                lNewSheet.Cells[1, col++] = string.Format("Regulon_{0}", c + 1);

            all.Columns.AutoFit();

            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;

        }


        private (List<string>,List<double>) GetOperonGenesFC(/*string operon,*/ string opid, List<BsuRegulons> lLst)
        {


            //SysData.DataRow[] lquery = gRefOperons.Select(string.Format("Operon = '{0}'", operon));
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

           
//            int nrHits = lquery.Count();

//            int __op_id = -1;
//            foreach (DataRow row in lquery)
//            {
//                int _op_id;
//                Int32.TryParse(row["op_id"].ToString(), out _op_id);

//                List<string> _genes = new List<string>();
//                List<double> _lfcs = new List<double>();

//                if (_op_id != __op_id)
//                {
//                    __op_id = _op_id;
                  
//                    string lgene = row["gene"].ToString();
//                    _genes.Add(lgene);

//                    BsuRegulons result = lLst.Find(item => item.GENE == lgene);
//                    if (result != null)
//                        _lfcs.Add(result.FC);
//                    else
//                        _lfcs.Add(Double.NaN);
//                }

//                lgenes.Add(_genes);
//                lFCs.Add(_lfcs);

////                if (lgene == operon && nrHits>1)
////                    continue;                                              
//            }   

                                   
            return (_genes,_lfcs);
        }

        private SysData.DataTable CreateOperonTable(SysData.DataTable aUsageTbl, List<BsuRegulons> lLst)
        {
            SysData.DataTable lTable = new SysData.DataTable();

            #region newoutput
            SysData.DataColumn col = new SysData.DataColumn("BSU", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("P-value", Type.GetType("System.Double"));
            lTable.Columns.Add(col);
            #endregion

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

            for(int nr=0;nr<maxGenesPerOperon;nr++)
            {
                col = new SysData.DataColumn(string.Format("gene_{0}",nr+1), Type.GetType("System.Double"));
                lTable.Columns.Add(col);
            }


            int noperons = 0;

            //SysData.DataTable uOperons = GetDistinctRecords(gRefOperons, new string[] { "Operon" });
            double lowVal = Properties.Settings.Default.fcLOW;
            for (int r = 0; r < lLst.Count; r++)
            {
                                
                string geneName = lLst[r].GENE;
                double lFC = lLst[r].FC;
                double lPval = lLst[r].PVALUE;


#if false   // accept all items
                bool accept = Properties.Settings.Default.use_pvalues ? lLst[r].PVALUE < Properties.Settings.Default.pvalue_cutoff : Math.Abs(lLst[r].FC) > lowVal;
                if (!accept)
                    continue;
#endif

                List<string> luOperons = new List<string>();

                
                // multiple operons for a single gene
                SysData.DataRow[] lOperons = gRefOperons.Select(string.Format("gene='{0}'", geneName));


                string operon = "";
                List<string> lgenes = new List<string>();
                List<double> lFCs = new List<double>();
                string opgenes = "";
                List<List<string>> llgenes = new List<List<string>>();

                int _m = 0;
                int _maxm = _m;
                foreach (DataRow row in lOperons)
                {
                    //if (operon.Length > 0)
                    //    operon += "#" + row["operon"].ToString();
                    //else
                    //    operon = row["operon"].ToString();
                    luOperons.Add(row["operon"].ToString());
                    //operon = row["operon"].ToString();
                    //operon = row["operon"].ToString();

                    // multiple genes per operon
                    //(List<List<string>> _lgenes, List<List<double>>  _lFCs) = GetOperonGenesFC(row["operon"].ToString(), row["op_id"], lLst);
                    (List<string> _lgenes, List<double> _lFCs) = GetOperonGenesFC(row["op_id"].ToString(), lLst);
                    llgenes.Add(_lgenes);


                    if (_lgenes.Count > lgenes.Count)
                    {
                        _maxm = _m;
                        operon = row["operon"].ToString();
                        lgenes = new List<string>(_lgenes);
                        lFCs = new List<double>(_lFCs);
                    }

                    _m++;

                }



                noperons = luOperons.Count();

                //luOperons.Sort((x, y) => x.Length.CompareTo(y.Length));

                //luOperons.Select(x=>x.Length>luOperons[id].Length)

                if(operon.Length>0)
                //foreach(string operon in luOperons)
                {
                    operon = luOperons[_maxm];
                    lgenes = llgenes[_maxm];
                    opgenes = string.Join("-", lgenes.ToArray());
                    llgenes.Remove(lgenes);

                    foreach (List<string> _item in llgenes)
                    {
                        opgenes = opgenes + Environment.NewLine + string.Join("-", _item.ToArray());
                    }



                    //(List<string> lgenes, List<double> lFCs) = GetOperonGenesFC(luOperons[0], lLst);
                    int nrgenes = lgenes.Count;
                    //string opgenes = string.Join("-", lgenes.ToArray());
                    // addrow
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["FC"] = lLst[r].FC;
                    lRow["P-Value"] = lLst[r].PVALUE;

                    lRow["gene"] = geneName;
                    lRow["operon"] = operon;
                    lRow["nroperons"] = noperons;
                    lRow["nrgenes"] = nrgenes;
                    lRow["operon_genes"] = opgenes;
                    for(int i=0;i<nrgenes;i++)
                    {
                      if(!(lFCs[i] is Double.NaN))
                            lRow[string.Format("gene_{0}", i + 1)] = lFCs[i];
                    }
                }
                else
                {
                    
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["FC"] = lLst[r].FC;
                    lRow["P-Value"] = lLst[r].PVALUE;
                    
                }
                                
            }


            return lTable;
        }

        private SysData.DataTable CreateCombinedTable(SysData.DataTable aUsageTbl, List <BsuRegulons> lLst)
        {
            SysData.DataTable lTable = new SysData.DataTable();

            
            SysData.DataColumn col = new SysData.DataColumn("BSU", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("GENE", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            lTable.Columns.Add(col);


            col = new SysData.DataColumn("PVALUE", Type.GetType("System.Double"));
            lTable.Columns.Add(col);
            
            //if (gOperonOutput)
            //{
            //    col = new SysData.DataColumn("OPERON", Type.GetType("System.String"));
            //    lTable.Columns.Add(col);
            //}

            int maxRegulons = 0;
            for (int i=0;i< lLst.Count; i++)
            {
                if (maxRegulons < lLst[i].REGULONS.Count)
                    maxRegulons = lLst[i].REGULONS.Count;
            }
            
            for(int i=0;i<maxRegulons;i++)
            {
                col = new SysData.DataColumn(string.Format("Regulon_{0}",i+1), Type.GetType("System.String"));
                lTable.Columns.Add(col);

            }
            
            double lowVal = Properties.Settings.Default.fcLOW;

            for (int r = 0; r < lLst.Count; r++)
            {
                // continue depending on value of lowest fc definition
                bool accept = Properties.Settings.Default.use_pvalues ? lLst[r].PVALUE < Properties.Settings.Default.pvalue_cutoff : Math.Abs(lLst[r].FC) > lowVal;

                if (accept)
                {
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["FC"] = lLst[r].FC;
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["GENE"] = lLst[r].GENE;
                    lRow["PVALUE"] = lLst[r].PVALUE;

                    //if (gRefOperons != null)
                    //{
                    //    string lOperon = "";
                    //    if (lLst[r].GENE != "")
                    //    {
                    //        SysData.DataRow[] lOperons = gRefOperons.Select(string.Format("gene='{0}'", lLst[r].GENE));
                    //        List<string> strOperons = new List<string>();
                    //        for (int i = 0; i < lOperons.Length; i++)
                    //        {
                    //            strOperons.Add(lOperons[i]["operon"].ToString());
                    //        }

                    //        lOperon = String.Join(", ", strOperons.ToArray());
                    //    }

                    //    lRow["OPERON"] = lOperon;
                    //}

                    double FC = lLst[r].FC;

                    for (int i = 0; i < lLst[r].REGULONS.Count; i++)
                    {

                        // check association direction 
                        bool posAssoc = lLst[r].UP.Contains(i) ? true : false;
                        // depending on the association in the table the cell color is red or green
                        char clrC = posAssoc ? '&' : '!';

                        SysData.DataRow[] lHit = aUsageTbl.Select(string.Format("Regulon = '{0}'", lLst[r].REGULONS[i]));
                        double nrUP = Double.Parse(lHit[0]["nr_UP"].ToString());
                        double nrDOWN = Double.Parse(lHit[0]["nr_DOWN"].ToString());
                        Double.TryParse(lHit[0]["perc_UP"].ToString(),out double percUP);
                        Double.TryParse(lHit[0]["perc_DOWN"].ToString(),out double percDOWN);

                        double percRel = Double.Parse(lHit[0]["totrelperc"].ToString());

                        string lVal = "";

                        // logical association
                        if ((posAssoc && FC > 0)||(!posAssoc && FC<0))
                        {
                            if (nrUP > nrDOWN)
                                lVal = percUP.ToString("P0") + "@"+ clrC + percRel.ToString("P0") + "-tot";
                            if (nrDOWN > nrUP)
                                lVal = percDOWN.ToString("P0") + "#" +clrC + percRel.ToString("P0") + "-tot";
                        }
                        if (nrUP == nrDOWN)
                            lVal = "0%-" + percRel.ToString("P0") + "-tot";
                        
                        // false postive/negative
                        if ((posAssoc && FC < 0) || (!posAssoc && FC > 0))
                        {
                            if (nrUP > nrDOWN)
                            {                                
                                if(percUP < 1.0)
                                    lVal = (1.0 - percUP).ToString("P0") + "#" +clrC + percRel.ToString("P0") + "-tot";
                                else
                                    lVal = percUP.ToString("P0") + "@" + clrC + percRel.ToString("P0") + "-tot";
                            }
                            if (nrDOWN > nrUP)
                            {
                                if(percDOWN<1.0)
                                    lVal = (1.0 - percDOWN).ToString("P0") + "@" + clrC + percRel.ToString("P0") + "-tot";
                                else
                                    lVal = percDOWN.ToString("P0") + "#"+clrC + percRel.ToString("P0") + "-tot";
                            }
                        }                               
                            
                        lRow[string.Format("Regulon_{0}", i + 1)] = lLst[r].REGULONS[i] + " " + lVal;

                    }
                }
            }


            for (int i = maxRegulons; i > 0 ;i--)
            {
                string columnName = string.Format("Regulon_{0}", i);
                object lRes = lTable.Compute(string.Format("COUNT({0})", columnName),""); 
                int lCount = Int16.Parse(lRes.ToString());
                if(lCount==0)                    
                    lTable.Columns.Remove(columnName);
            }

            return lTable;
        }

        private (SysData.DataTable, SysData.DataTable) CreateUsageTable(List<FC_BSU> aList)
        {
            {
                SysData.DataTable _fc_BSU = ReformatResults(aList);

                SysData.DataTable lTable = new SysData.DataTable();
                SysData.DataTable lTableCombine = new SysData.DataTable(); // table for combined summary


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

                col = new SysData.DataColumn("perc_DOWN", Type.GetType("System.Double"));
                lTable.Columns.Add(col);
                col = new SysData.DataColumn("perc_UP", Type.GetType("System.Double"));
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


                col = new SysData.DataColumn("nr_DOWN", Type.GetType("System.Double"));
                lTableCombine.Columns.Add(col);
                col = new SysData.DataColumn("nr_UP", Type.GetType("System.Double"));
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

                    (int nrUP, int nrDOWN, int nrTOT) = CalculateFPRatio(_tmp);

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

                    SysData.DataRow lNewRow = lTable.Rows.Add();

                    lNewRow["CountData"] = _tmp.Length;
                    lNewRow["Count"] = _tmp2[0]["Count"];

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

                    lNewRow["totrel"] = up2 + up3 + up4 + down2 + down3 + down4; // nrTOT;


                    // nrUP and nrDOWN contain the counts of those genes that were defined as up or down regulated that had a 'significant' fc.
                    // this was, false positive can be identified
                    if (nrTOT > 0)
                    {
                        lNewRow["perc_DOWN"] = (double)nrDOWN / (double)(nrTOT);
                        lNewRow["perc_UP"] = (double)nrUP / (double)(nrTOT);
                    }


                    double lCount = (double)_tmp.Length;

                    lNewRow = lTableCombine.Rows.Add();


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
                        lNewRow["perc_DOWN"] = (double)nrDOWN / (double)(nrTOT);
                        lNewRow["perc_UP"] = (double)nrUP / (double)(nrTOT);
                    }

                    lNewRow["nr_DOWN"] = ((double)nrDOWN) / lCount;
                    lNewRow["nr_UP"] = ((double)nrUP) / lCount;


                }

                SysData.DataView dv = lTable.DefaultView;
                dv.Sort = "totrel desc";

                return (dv.ToTable(), lTableCombine);
            }
        }
       

        private RibbonDropDownItem getItemByValue(RibbonDropDown ctrl, string value)
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

        private void LoadDirectionOptions()
        {
            SysData.DataView view = new SysData.DataView(gRefWB);
            SysData.DataTable distinctValues = view.ToTable(true, Properties.Settings.Default.referenceDIR);

            foreach (SysData.DataRow row in distinctValues.Rows)
            {
                gAvailItems.Add(row.ItemArray[0].ToString());
            }
        }

        private void load_Worksheets()
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



        private void load_OperonSheet()
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

        private void load_CatFile()
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

            RibbonDropDownItem ddItem = getItemByValue(ddBSU, Properties.Settings.Default.referenceBSU);
            if (ddItem != null)
                ddBSU.SelectedItem = ddItem;

            ddItem = getItemByValue(ddRegulon, Properties.Settings.Default.referenceRegulon);
            if (ddItem != null)
                ddRegulon.SelectedItem = ddItem;

            ddItem = getItemByValue(ddDir, Properties.Settings.Default.referenceDIR);
            if (ddItem != null)
                ddDir.SelectedItem = ddItem;

            ddItem = getItemByValue(ddGene, Properties.Settings.Default.referenceGene);
            if (ddItem != null)
                ddGene.SelectedItem = ddItem;

            ddBSU.Enabled = true;
            ddRegulon.Enabled = true;
            ddDir.Enabled = true;
            ddGene.Enabled = true;
            btRegDirMap.Enabled = true;
            //edtMaxGroups.Enabled = true;
            //btnPalette.Enabled = true;

            gApplication.EnableEvents = true;


        }


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

        private void btLoad_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;
            if (LoadData())
            {
                gOperonOutput = LoadOperonData();
                gCatOutput = LoadCategoryData();
               // cbUseCategories.Enabled = gCatOutput;                
                Fill_DropDownBoxes();
                if (gDownItems.Count == 0 && gUpItems.Count == 0 && gAvailItems.Count == 0)
                    LoadDirectionOptions();
                //btApply.Enabled = true;
                //btPlot.Enabled = true; 
                btnSelect.Enabled = true;
                toggleButton1.Enabled = true;
                LoadFCDefaults();
                ResetTables();
                //EnableOutputOptions(true);
            }
            gApplication.EnableEvents = true;
        }

        private void LoadFCDefaults()
        {
            ebLow.Text = Properties.Settings.Default.fcLOW.ToString();
            ebMid.Text = Properties.Settings.Default.fcMID.ToString();
            ebHigh.Text = Properties.Settings.Default.fcHIGH.ToString();
            editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();
        }

        private void EnableItems(bool enable)
        {
            btLoad.Enabled = enable;
            ddBSU.Enabled = enable;
            ddRegulon.Enabled = enable;
            ddGene.Enabled = enable;
            //btPlot.Enabled = enable;
            //edtMaxGroups.Enabled = enable;
            //btnPalette.Enabled = enable;

        }
    
        private void ddBSU_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceBSU = ddBSU.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
        }

        private void ddRegulon_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceRegulon = ddRegulon.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
        }

        private void btRegDirMap_Click(object sender, RibbonControlEventArgs e)
        {
            dlgUpDown dlgUD = new dlgUpDown(gAvailItems, gUpItems, gDownItems);
            dlgUD.ShowDialog();

            storeValue("directionMapUnassigned", gAvailItems);
            storeValue("directionMapUp", gUpItems);
            storeValue("directionMapDown", gDownItems);

            SetFlags(UPDATE_FLAGS.ALL);            

        }

        private void ddDir_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceDIR = ddDir.SelectedItem.Label;
            gAvailItems.Clear();
            gUpItems.Clear();
            gDownItems.Clear();
            LoadDirectionOptions();
            SetFlags(UPDATE_FLAGS.ALL);
        }

        private void validateTextBoxData(RibbonEditBox bx)
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

        private void ebLow_TextChanged(object sender, RibbonControlEventArgs e)
        {
            validateTextBoxData(ebLow);
        }

        private void ebMid_TextChanged(object sender, RibbonControlEventArgs e)
        {
            validateTextBoxData(ebMid);
        }

        private void ebHigh_TextChanged(object sender, RibbonControlEventArgs e)
        {
            validateTextBoxData(ebHigh);
        }

#region main_routines
        private void btPlot_Click(object sender, RibbonControlEventArgs e)
        {
            if(!(Properties.Settings.Default.catPlot ||Properties.Settings.Default.regPlot  || Properties.Settings.Default.distPlot))
            {
                MessageBox.Show("Please select at least one plot to generate");
                return;
            }


            if ((Properties.Settings.Default.catPlot || Properties.Settings.Default.regPlot)) //& gNeedsUpdate.Check(UPDATE_FLAGS.PCat))
            {
                dlgTreeView dlg = new dlgTreeView(categoryView: cbUseCategories.Checked, spreadingOptions: Properties.Settings.Default.catPlot,rankingOptions: Properties.Settings.Default.regPlot);

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
                        (gOutput, gList) = GenerateOutput(suppressOutput: true);

                        if (gOutput != null && gList != null)
                        {
                            UnSetFlags(UPDATE_FLAGS.TMapped);
                            (gSummary, gCombineInfo) = CreateUsageTable(gOutput);
                            UnSetFlags(UPDATE_FLAGS.TCombined);
                            //UnSetFlags(UPDATE_FLAGS.TSummary);
                            //gNeedsUpdate.UnSet(UPDATE_FLAGS.TCombined);
                        }
                        
                        //gNeedsUpdating = false;
                    }

                    //if (gEnrichmentAnalysis == null)
                    //gEnrichmentAnalysis = new EnrichmentAnalysis(gApplication);

                    if ((gOutput != null && gSummary != null && dlg.GetSelection().Count() > 0) ) //&& gNeedsUpdate.Check(UPDATE_FLAGS.PCat))
                    {
                        if (Properties.Settings.Default.catPlot)
                        {
                            //CategoryPlot(gOutput, gSummary, dlg.GetSelection(),outputTable:dlg.selectTableOutput());
                            SpreadingPlot(gOutput, gSummary, dlg.GetSelection(),topTenFC:dlg.getTopFC(),topTenP:dlg.getTopP(), outputTable: dlg.selectTableOutput());


                            //UnSetFlags(UPDATE_FLAGS.PCat);
                            //gNeedsUpdate.UnSet(UPDATE_FLAGS.PCat);
                        }
                    //}
                    //if ((gOutput != null && gSummary != null && dlg.GetSelection().Count() > 0) ) //&& gNeedsUpdate.Check(UPDATE_FLAGS.PRegulon))
                    //{ 
                        if (Properties.Settings.Default.regPlot)
                        {
                            RankingPlot(gOutput, gSummary, dlg.GetSelection(),tableOutput:dlg.selectTableOutput(),splitNP:dlg.GetSplitOption()); // wordt optie in category plot hierboven
                            //private void CreateRegulonPlotDataSheet(List<element_fc> theElements)

                            // data for now.. to be changed in plot

                            //UnSetFlags(UPDATE_FLAGS.PRegulon);
                            //gNeedsUpdate.UnSet(UPDATE_FLAGS.PRegulon);
                        }
                    }
                }

            }


            // verschil tussen tabel maken en data updaten!!!! 1-2-2021

            if(Properties.Settings.Default.distPlot)
            {
                if (gOutput == null || gSummary == null || gNeedsUpdate.Check(UPDATE_FLAGS.TMapped))
                {
                    (gOutput, gList) = GenerateOutput(suppressOutput: true);
                    if (gOutput != null && gList != null)
                    {
                        UnSetFlags(UPDATE_FLAGS.TMapped);                        
                        (gSummary, gCombineInfo) = CreateUsageTable(gOutput);
                        UnSetFlags(UPDATE_FLAGS.TCombined);                        
                    }                    
                }

                if ((gOutput != null && gSummary != null) ) //&& gNeedsUpdate.Check(UPDATE_FLAGS.PDist))
                {
                    DistributionPlot(gOutput, gSummary);
                    //UnSetFlags(UPDATE_FLAGS.PDist);
                }
            }


        }
        
        private void SetFlags(UPDATE_FLAGS f)
        {
            gNeedsUpdate = (byte) (gNeedsUpdate | (byte)f);
        }

        private void UnSetFlags(UPDATE_FLAGS f)
        {
            gNeedsUpdate = (byte)(gNeedsUpdate & (byte)~f);
        }

        private bool NeedsUpdate(UPDATE_FLAGS f)
        {
            return gNeedsUpdate.Check(f);
        }

        private bool AnyUpdate()
        {
            return gNeedsUpdate.Any();
        }

        private bool NoUpdate()
        {
            return gNeedsUpdate.None();
        }


        private void btApply_Click(object sender, RibbonControlEventArgs e)
        {


            if( !(Properties.Settings.Default.tblMap || Properties.Settings.Default.tblSummary || Properties.Settings.Default.tblCombine || Properties.Settings.Default.tblOperon))
            {
                MessageBox.Show("Please select at least one output table to generate");
                return;
            }
            
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            if (NoUpdate())
                return;
            
            //if ( (gOutput == null || gList == null) || NeedsUpdate(UPDATE_FLAGS.TMapped))
            //{
            //    (gOutput, gList) = GenerateOutput();

            //    if (gOutput is null || gList is null)
            //        return;

            //    UnSetFlags(UPDATE_FLAGS.TMapped);
            //}
            if(Properties.Settings.Default.tblMap)
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
                SysData.DataTable lCombined = CreateCombinedTable(gCombineInfo, gList);
                CreateCombinedSheet(lCombined);
                UnSetFlags(UPDATE_FLAGS.TCombined);                
              
            }            
                

            if ( Properties.Settings.Default.tblOperon && gOperonOutput) // can combine table/sheet because it's a quick routine
            {
                SysData.DataTable tblOperon = CreateOperonTable(gCombineInfo, gList);
                CreateOperonSheet(tblOperon);
                UnSetFlags(UPDATE_FLAGS.TOperon);                
            }
            
            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }



        private element_fc CatElements2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements, int topTenFC=-1, int topTenP=-1)
        {

            //List<element_fc> element_Fcs = new List<element_fc>();
            element_fc element_Fcs = new element_fc();
            SysData.DataView dataViewCat = gCategories.AsDataView();


            List<summaryInfo> _All = new List<summaryInfo>();
            List<summaryInfo> _Pos = new List<summaryInfo>();
            List<summaryInfo> _Neg = new List<summaryInfo>();
            List<summaryInfo> _Com = new List<summaryInfo>();

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

#region toberemoved


                //element_Fc.sd = _fcsP.Count>0?_fcsP.sd():0;

                //element_Fc.madFC_P = _fcsP.Count>0?_fcsP.mad():0;                    
                //element_Fc.madFC_N = _fcsN.Count>0?_fcsN.mad():0;
                //element_Fc.madFC_T = _fcsT.Count>0?_fcsT.mad():0;

#endregion

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
            element_Fcs.Pos = _Pos;
            element_Fcs.Neg = _Neg;


            if (Properties.Settings.Default.useSort)
            {
                double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();

                //List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                //return (element_Fcs, sortedIndex);

                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList();

                if (Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if (topTenFC > 0)
            {
                double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                //var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => Math.Abs(x.Key)).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).Where(x => x.fc_values.Length > 0).ToList();
                element_Fcs.All = element_Fcs.All.GetRange(0, topTenFC);

                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if (topTenP > 0)// too lazy.. just copy here
            {
                // assertion... it's -10 log(p) -> so the higher, the better
                double[] __values = element_Fcs.All.Select(x => -Math.Log(x.p_average)).ToArray();
                var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).Where(x => x.fc_values.Length > 0).ToList();
                element_Fcs.All = element_Fcs.All.GetRange(0, topTenP);

                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }

            return element_Fcs;
        }


        private element_fc Regulons2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements, int topTenFC = -1, int topTenP = -1, bool splitNP=false) // HashSet<string> regulons)
        {
            element_fc element_Fcs = new element_fc();
            
            List<summaryInfo> _All = new List<summaryInfo>();
            List<summaryInfo> _Pos = new List<summaryInfo>();
            List<summaryInfo> _Neg = new List<summaryInfo>();
            List<summaryInfo> _Com = new List<summaryInfo>();

            foreach (cat_elements el in cat_Elements)
            {                
                dataView.RowFilter = String.Format("Regulon='{0}'", el.catName);                              

                SysData.DataTable _dataTable = dataView.ToTable();
                //element_Fc.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);

                // find genes for the regulon/category

                summaryInfo __All = new summaryInfo();
                summaryInfo __Pos = new summaryInfo();
                summaryInfo __Neg = new summaryInfo();
                //summaryInfo __Com = new summaryInfo();

                __All.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                __Pos.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);
                __Neg.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);


                List<double> _fcsA = new List<double>();
                List<string> _genesA = new List<string>();
                List<double> _pvaluesA = new List<double>();


                List<double> _fcsP = new List<double>();
                List<string> _genesP = new List<string>();
                List<double> _pvaluesP = new List<double>();

                List<double> _fcsN = new List<double>();
                List<string> _genesN = new List<string>();
                List<double> _pvaluesN = new List<double>();


                if (_dataTable.Rows.Count > 0)
                {

                    //__Com.catName = string.Format("{0} ({1})", el.catName, _dataTable.Rows.Count);



                    for (int i = 0; i < _dataTable.Rows.Count; i++)
                    {

                        double fc = (double)_dataTable.Rows[i]["FC"];

                        _genesA.Add(_dataTable.Rows[i]["Gene"].ToString());
                        _fcsA.Add(fc);
                        _pvaluesA.Add(double.Parse(_dataTable.Rows[i]["Pvalue"].ToString()));

                        if (fc >= 0)
                        {
                            _genesP.Add(_dataTable.Rows[i]["Gene"].ToString());
                            _fcsP.Add(fc);
                            _pvaluesP.Add(double.Parse(_dataTable.Rows[i]["Pvalue"].ToString()));
                        }
                        else if (fc < 0)
                        {
                            _genesN.Add(_dataTable.Rows[i]["Gene"].ToString());
                            _fcsN.Add(fc);
                            _pvaluesN.Add(double.Parse(_dataTable.Rows[i]["Pvalue"].ToString()));
                        }
                    }
                }



                __Pos.fc_average = _fcsP.Count>0?_fcsP.Average() : Double.NaN;                    
                __Neg.fc_average = _fcsN.Count>0?_fcsN.Average() : Double.NaN;
                __All.fc_average = _fcsA.Count>0?_fcsA.Average() : Double.NaN;

                __Pos.fc_values = _fcsP.Count > 0 ? _fcsP.ToArray() : new double[0];// { 0 };
                __Neg.fc_values = _fcsN.Count > 0 ? _fcsN.ToArray() : new double[0];// { 0 };
                __All.fc_values = _fcsA.Count > 0 ? _fcsA.ToArray() : new double[0];// { 0 };

                __Pos.fc_mad = _fcsP.Count > 0 ? _fcsP.mad() : Double.NaN;
                __Neg.fc_mad = _fcsN.Count > 0 ? _fcsN.mad() : Double.NaN;
                __All.fc_mad = _fcsA.Count > 0 ? _fcsA.mad() : Double.NaN;

                #region toberemoved


                //element_Fc.sd = _fcsP.Count>0?_fcsP.sd():0;

                //element_Fc.madFC_P = _fcsP.Count>0?_fcsP.mad():0;                    
                //element_Fc.madFC_N = _fcsN.Count>0?_fcsN.mad():0;
                //element_Fc.madFC_T = _fcsT.Count>0?_fcsT.mad():0;

                #endregion

                __Pos.genes = _genesP.Count > 0 ? _genesP.ToArray() : new string[0]; // { "" };
                __Neg.genes = _genesN.Count > 0 ? _genesN.ToArray() : new string[0];// { "" };
                __All.genes = _genesA.Count > 0 ? _genesA.ToArray() : new string[0];// { "" };


                __Pos.p_values = _pvaluesP.Count > 0 ? _pvaluesP.ToArray() : new double[0];// { };
                __Neg.p_values = _pvaluesN.Count > 0 ? _pvaluesN.ToArray() : new double[0];// { };
                __All.p_values = _pvaluesA.Count > 0 ? _pvaluesA.ToArray() : new double[0];// { };

                __Pos.p_average = _pvaluesP.Count>0 ? _pvaluesP.paverage() : Double.NaN;
                __Neg.p_average = _pvaluesN.Count>0 ? _pvaluesN.paverage() : Double.NaN;
                __All.p_average = _pvaluesA.Count>0 ? _pvaluesA.paverage() : Double.NaN;

                __Pos.p_mad = _pvaluesP.Count>0?_pvaluesP.mad() : Double.NaN;
                __Neg.p_mad = _pvaluesN.Count>0?_pvaluesN.mad() : Double.NaN;
                __All.p_mad = _pvaluesA.Count>0?_pvaluesA.mad() : Double.NaN;

                _All.Add(__All);
                _Pos.Add(__Pos);
                _Neg.Add(__Neg);


                element_Fcs.All = _All;
                element_Fcs.Pos = _Pos;
                element_Fcs.Neg = _Neg;
                 
            }


            if (Properties.Settings.Default.useSort)
            {
                double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();

                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                //return (element_Fcs, sortedIndex);
                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList();

                if (Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if (topTenFC > 0 )
            {
                double[] __values = element_Fcs.All.Select(x => x.fc_average).ToArray();
                //var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => Math.Abs(x.Key)).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => Math.Abs(x.Key)).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();                
                // remove elements with no genes associated
                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList().Where(x=> x.fc_values.Length > 0).ToList();
                element_Fcs.All = element_Fcs.All.GetRange(0, topTenFC);
                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }
            else if(topTenP>0 )// too lazy.. just copy here
            {
                // assertion... it's -10 log(p) -> so the higher, the better
                double[] __values = element_Fcs.All.Select(x =>-Math.Log(x.p_average)).ToArray();
                //var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x =>x.Key).ToList();                
                var sortedElements = __values.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderByDescending(x => x.Key).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();                
                // remove elements with no genes associated
                element_Fcs.All = sortedElements.Select(x => element_Fcs.All[x.Value]).ToList().Where(x => x.fc_values.Length > 0).ToList();
                element_Fcs.All = element_Fcs.All.GetRange(0, topTenP);
                if (!Properties.Settings.Default.sortAscending)
                    element_Fcs.All.Reverse();
            }                                        
       
       
            return element_Fcs;
        }

    
        // output of all genes in table
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

#region toberemoved
        //public Excel.Chart MyTmpPlot(List<element_fc> element_Fcs)
        //{
            

        //    if (gApplication == null)
        //        return null;

        //    Excel.Worksheet aSheet = gApplication.Worksheets.Add();

            
        //    //var missing = System.Type.Missing;

        //    Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
        //    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
        //    Excel.Chart chartPage = myChart.Chart;            

        //    chartPage.ChartType = Excel.XlChartType.xlXYScatter;

        //    var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

        //    int nrCategories = element_Fcs.Count;

        //    double MMAX = 0;
        //    double MMIN = 0;

        //    for (int _i = 0; _i < nrCategories; _i++)
        //    {
        //        if (element_Fcs[_i].fcP != null)
        //        {
        //            if (element_Fcs[_i].fcP.Min() < MMIN)
        //                MMIN = element_Fcs[_i].fcP.Min();
        //            if (element_Fcs[_i].fcP.Max() > MMAX)
        //                MMAX = element_Fcs[_i].fcP.Max();
        //        }
        //    }


        //    foreach (var element_Fc in element_Fcs.Select((value, index) => new { value, index }))
        //    {
        //        var xy1 = series.NewSeries();
        //        xy1.Name = element_Fc.value.catName;
        //        xy1.ChartType = Excel.XlChartType.xlXYScatter;
        //        if (element_Fc.value.fcP != null)
        //        {
        //            xy1.XValues = element_Fc.value.fcP;
        //            xy1.Values = Enumerable.Repeat(element_Fc.index + 0.5, element_Fc.value.fcP.Length).ToArray();
        //            xy1.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
        //            xy1.MarkerSize = 2;
        //            xy1.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, 0.1);
        //            Excel.ErrorBars errorBars = xy1.ErrorBars;
        //            errorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap;
        //            errorBars.Format.Line.Weight = 1.25f;

        //            // give each serie different color
        //            switch (element_Fc.index % 6)
        //            {
        //                case 0:
        //                    errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent1;
        //                    break;
        //                case 1:
        //                    errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent2;
        //                    break;
        //                case 2:
        //                    errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent3;
        //                    break;
        //                case 3:
        //                    errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent4;
        //                    break;
        //                case 4:
        //                    errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
        //                    break;
        //                case 5:
        //                    errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6;
        //                    break;
        //            }


        //        }
        //        var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
        //        //yAxis.AxisTitle.Text = "Regulon";
        //        Excel.TickLabels labels = yAxis.TickLabels;
        //        labels.Offset = 1;
        //    }




        //    chartPage.ChartColor = 1; // Properties.Settings.Default.defaultPalette;

        //    // as a last step, add the axis labels series

        //    if (true)
        //    {

        //        var xy2 = series.NewSeries();

        //        xy2.ChartType = Excel.XlChartType.xlXYScatter;
        //        //# Excel.Range rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i*2)+1], _tmpSheet.Cells[6, (i * 2) + 1]];
        //        xy2.XValues = Enumerable.Repeat(MMIN, nrCategories).ToArray();

        //        //rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i * 2) + 2], _tmpSheet.Cells[6, (i * 2) + 2]];
        //        double[] yv = new double[nrCategories];
        //        for (int _i = 0; _i < nrCategories; _i++)
        //        {
        //            yv[_i] = ((float)_i) + 0.5f;
        //        }

        //        xy2.Values = yv;

        //        xy2.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
        //        xy2.HasDataLabels = true;

        //        for (int _i = 0; _i < nrCategories; _i++)
        //        {
        //            xy2.DataLabels(_i + 1).Text = element_Fcs[_i].catName;
        //        }

        //        xy2.DataLabels().Position = Excel.XlDataLabelPosition.xlLabelPositionLeft;

        //    }


        //    chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
        //    chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();            
        //    chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
        //    chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
        //    chartPage.Axes(Excel.XlAxisType.xlValue).MaximumScale = nrCategories;
        //    chartPage.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
            
        //    chartPage.Legend.Delete();
        //    chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);
          
        //    aSheet.Delete();
        //    return chartPage;

        //}
#endregion

        private void DistributionPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            SysData.DataTable _fc_BSU_ = ReformatResults(aOutput);
            SysData.DataTable _fc_BSU = GetDistinctRecords(_fc_BSU_, new string[] { "Gene","FC"});

            (List<double> sFC, List<int> sIdx) = SortedFoldChanges(_fc_BSU);

            int chartNr = nextWorksheet("DistributionPlot_");
            string chartName = "DistributionPlot_" + chartNr.ToString();

            Excel.Chart aChart = PlotRoutines.CreateDistributionPlot(sFC,sIdx, chartName);
            this.RibbonUI.ActivateTab("TabGINtool");


            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }


        private void SpreadingPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements, int topTenFC=-1,int topTenP=-1,bool outputTable = false )
        {
            //gApplication.EnableEvents = false;
            //gApplication.DisplayAlerts = false;

            AddTask(TASKS.CATEGORY_CHART);
                       
            SysData.DataTable _fc_BSU = ReformatResults(aOutput);

            cat_Elements = GetUniqueElements(cat_Elements);

            List<element_fc> element_Fcs = new List<element_fc>();

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
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements, topTenFC, topTenP);
           

            
            string postFix = topTenFC>-1 ? string.Format("Top{0}FC",topTenFC) : (topTenP > -1 ? string.Format("Top{0}P", topTenP) : "");

            
            string chartBase = (Properties.Settings.Default.useCat ? string.Format("CatSpreadPlot{0}_", postFix) : string.Format("RegSpreadPlot{0}_", postFix));

            int chartNr = nextWorksheet(chartBase);


            string chartName = chartBase + chartNr.ToString();
            



            Excel.Chart aChart = PlotRoutines.CreateCategoryPlot(catPlotData, chartName);

            if (outputTable)
            {
                catPlotData.All.Reverse();
                CreateExtendedRegulonCategoryDataSheet(catPlotData, chartName);
            }

#if CLICK_CHART
            
            Excel.Chart aChart = gApplication.ActiveChart;
            aChart.MouseDown += new Excel.ChartEvents_MouseDownEventHandler(AChart_MouseDown);
            gCharts.Add(new chart_info(aChart, catPlotData));
#endif

            this.RibbonUI.ActivateTab("TabGINtool");

            RemoveTask(TASKS.CATEGORY_CHART);

            //gApplication.EnableEvents = true;
            //gApplication.DisplayAlerts = true;            
        }

#if CLICK_CHART

        private void AChart_MouseDown(int Button, int Shift, int x, int y)
        {
            var aChart = gApplication.ActiveChart;
            chart_info cI = gCharts.isFound(aChart);
            if(!cI.Equals(ClassExtensions.Empty))
            {
                System.Console.WriteLine("yes");
                int elementId = 0;
                int arg1 = 0, arg2 = 0;
                cI.chart.GetChartElement(x, y, ref elementId, ref arg1, ref arg2);
            }

        }
#endif


        private void AChart_MouseMove(int Button, int Shift, int x, int y)
        {
            System.Console.WriteLine("Hello World");
                
                //throw new NotImplementedException();
        }


        //private void SpreadingSheet(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements)
        //{

        //    AddTask(TASKS.REGULON_CHART);

        //    SysData.DataTable _fc_BSU = ReformatResults(aOutput);


        //    List<element_fc> element_Fcs = new List<element_fc>();

        //    // HashSet ensures unique list
        //    HashSet<string> lRegulons = new HashSet<string>();

        //    foreach (SysData.DataRow row in aSummary.Rows)
        //        lRegulons.Add(row.ItemArray[0].ToString());

        //    SysData.DataView dataView = _fc_BSU.AsDataView();
        //    List<element_fc> catPlotData = null;
        //    if (Properties.Settings.Default.useCat)
        //    {
        //        catPlotData = CatElements2ElementsFC(dataView, cat_Elements);
        //    }
        //    else
        //        catPlotData = Regulons2ElementsFC(dataView, cat_Elements);

        //    CreateExtendedRegulonCategoryDataSheet(catPlotData);            

        //    RemoveTask(TASKS.REGULON_CHART);

        //    //gApplication.EnableEvents = true;
        //    //gApplication.DisplayAlerts = true;
        //}


        private void RankingPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements, bool tableOutput=false, bool splitNP=false)
        {
            //gApplication.EnableEvents = false;
            //gApplication.DisplayAlerts = false;

            AddTask(TASKS.REGULON_CHART);

            SysData.DataTable _fc_BSU = ReformatResults(aOutput);



            cat_Elements = GetUniqueElements(cat_Elements);


            List<element_fc> element_Fcs = new List<element_fc>();

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
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements,splitNP:splitNP);

            //CreateExtendedRegulonCategoryDataSheet(catPlotData);
            (List<element_rank> plotData,List<summaryInfo> _all, List<summaryInfo> _pos, List < summaryInfo > _neg, List <summaryInfo> _com) = CreateRankingPlotData(catPlotData);


            int chartNr = Properties.Settings.Default.useCat ? nextWorksheet("CatRankPlot_") : nextWorksheet("RegRankPlot_");
            string chartName = (Properties.Settings.Default.useCat ? "CatRankPlot_" : "RegRankPlot_") + chartNr.ToString();
                        

            //if (tableOutput)
            Excel.Worksheet lRankingSheet = CreateRankingDataSheet(catPlotData, _all,_pos,_neg,_com);

            PlotRoutines.CreateRankingPlot(lRankingSheet, plotData, chartName);

            this.RibbonUI.ActivateTab("TabGINtool");

            RemoveTask(TASKS.REGULON_CHART);

            //gApplication.EnableEvents = true;
            //gApplication.DisplayAlerts = true;
        }

        
        private List<cat_elements> GetUniqueElements(List<cat_elements> elements)
        {
            List<cat_elements> result = new List<cat_elements>();
            foreach(cat_elements el in elements)
            {
                cat_elements elo = result.GetCatElement(el.catName);
                if (elo.catName==null)
                    result.Add(el);
                else
                {
                    System.Console.WriteLine("Hallo");
                }             

            }

            return result;
        }


        private string StripText(string str)
        {
            string name = str;
            string[] names = name.Split('(');
            int hit = names[0].ToUpper().IndexOf("REGULON");
            string newname = hit == -1 ? names[0] : names[0].Substring(0, hit);
            return newname;
        }


        private List<summaryInfo> SortedElements(List<summaryInfo> alist, SORTMODE mode = SORTMODE.FC, bool descending = true)
        {
            List<summaryInfo> _work = new List<summaryInfo>(alist.Where(x => x.fc_values.Length > 0));
            List<summaryInfo> _skip = new List<summaryInfo>(alist.Where(x => x.fc_values.Length == 0));
            

            if (mode == SORTMODE.CATNAME)
            {
                string[] __values = _work.Select(x => x.catName).ToArray();
                var sortedElements = descending ? __values.Select((x, i) => new KeyValuePair<string, int>(x, i)).OrderByDescending(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<string, int>(x, i)).OrderBy(x => x.Key).ToList();
                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();

                alist = sortedElements.Select(x => _work[x.Value]).ToList();
                alist.AddRange(_skip);

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
                alist.AddRange(_skip);

                return alist;
            }
        }

        private (List<element_rank>, List<summaryInfo> , List<summaryInfo>, List<summaryInfo>, List<summaryInfo>) CreateRankingPlotData(element_fc theElements)
        {

            //List<element_fc> toRemove = new List<element_fc>();
            //DataTable lTable = ElementsToTable(theElements);
            List<element_rank> element_Ranks = new List<element_rank>();
            
            // MAD values
            List<double> e1_m = new List<double>(), e2_m = new List<double>(), e3_m = new List<double>(), e4_m = new List<double>(), e5_m = new List<double>();
            // FC values
            List<double> e1_fc = new List<double>(), e2_fc = new List<double>(), e3_fc = new List<double>(), e4_fc = new List<double>(), e5_fc = new List<double>();
            // Counts
            List<int> e1_n = new List<int>(), e2_n = new List<int>(), e3_n = new List<int>(), e4_n = new List<int>(), e5_n = new List<int>();
            // CATEGORY/REGULON NAMES
            List<string> e1_s = new List<string>(), e2_s = new List<string>(), e3_s = new List<string>(), e4_s = new List<string>(), e5_s = new List<string>();

            

            foreach (summaryInfo sInfo in theElements.All)
            {
                List<double> _workfc = null;
                List<double> _workm = null;
                List<int> _workn = null;
                List<string> _works = null;

                if (sInfo.p_average <0.06125 && sInfo.genes[0]!="" )
                {
                    _workfc = e1_fc;
                    _workm = e1_m;
                    _workn = e1_n;
                    _works = e1_s;
                }

                if (sInfo.p_average >= 0.06125 && sInfo.p_average < 0.125 && sInfo.genes[0] != "")
                {
                    _workfc = e2_fc;
                    _workm = e2_m;
                    _workn = e2_n;
                    _works = e2_s;
                }


                if (sInfo.p_average >= 0.125 && sInfo.p_average < 0.25 && sInfo.genes[0] != "")
                {
                    _workfc = e3_fc;
                    _workm = e3_m;
                    _workn = e3_n;
                    _works = e3_s;
                }
                if (sInfo.p_average >= 0.25 && sInfo.p_average < 0.5 && sInfo.genes[0] != "")
                {
                    _workfc = e4_fc;
                    _workm = e4_m;
                    _workn = e4_n;
                    _works = e4_s;
                }


                if (sInfo.p_average >= 0.5 && sInfo.p_average <= 1 && sInfo.genes[0] != "")
                {
                    _workfc = e5_fc;
                    _workm = e5_m;
                    _workn = e5_n;
                    _works = e5_s;
                }

                if (_workfc != null)
                {

                    _workfc.Add(sInfo.fc_average);
                    _workm.Add(sInfo.fc_mad);
                    _workn.Add(sInfo.p_values != null ? sInfo.p_values.Length : 0);
                    _works.Add(StripText(sInfo.catName));
                }              
            }


            element_rank e1 = new element_rank(), e2 = new element_rank(), e3 = new element_rank(), e4 = new element_rank(), e5 = new element_rank();

            e1.catName = "p<0.0625";
            e1.average_fc = e1_fc.ToArray();
            e1.mad_fc = e1_m.ToArray();
            e1.nr_genes = e1_n.ToArray();
            e1.genes = e1_s.ToArray();
            
            e2.catName = "0.0625>=p<0.125";
            e2.average_fc = e2_fc.ToArray();
            e2.mad_fc = e2_m.ToArray();
            e2.nr_genes = e2_n.ToArray();
            e2.genes = e2_s.ToArray();

            e3.catName = "0.125>=p<0.25";
            e3.average_fc = e3_fc.ToArray();
            e3.mad_fc = e3_m.ToArray();
            e3.nr_genes = e3_n.ToArray();
            e3.genes = e3_s.ToArray();

            e4.catName = "0.25>=p<0.5";
            e4.average_fc = e4_fc.ToArray();
            e4.mad_fc = e4_m.ToArray();
            e4.nr_genes = e4_n.ToArray();
            e4.genes = e4_s.ToArray();

            e5.catName = "0.5>=p=<1";
            e5.average_fc = e5_fc.ToArray();
            e5.mad_fc = e5_m.ToArray();
            e5.nr_genes = e5_n.ToArray();
            e5.genes = e5_s.ToArray();

            element_Ranks.Add(e1);
            element_Ranks.Add(e2);
            element_Ranks.Add(e3);
            element_Ranks.Add(e4);
            element_Ranks.Add(e5);


            List<summaryInfo> all_elements = SortedElements(theElements.All,mode:SORTMODE.P,descending:false);
            List<summaryInfo> pos_elements = SortedElements(theElements.Pos);
            List<summaryInfo> neg_elements = SortedElements(theElements.Neg,descending:false);

            return (element_Ranks, all_elements, pos_elements,neg_elements,null);


        }

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
                
                for(int g=0;g< sInfo[r].genes.Count();g++)
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

      

        private DataTable UDElementsToTable(element_fc element_info)
        {

            List<summaryInfo> _tmp = SortedElements(element_info.All, mode: SORTMODE.CATNAME,descending:false);

            SysData.DataTable lTable = new SysData.DataTable("Elements");

            SysData.DataColumn regColumn = new SysData.DataColumn("Name", Type.GetType("System.String"));
            SysData.DataColumn cntColumnA = new SysData.DataColumn("CountT", Type.GetType("System.Int16"));            
            
            SysData.DataColumn statColumn1 = new SysData.DataColumn("Mode1", Type.GetType("System.String"));
            SysData.DataColumn cntColumn1 = new SysData.DataColumn("Count1", Type.GetType("System.Int16"));
            SysData.DataColumn percColumn1 = new SysData.DataColumn("Perc1", Type.GetType("System.Double"));
            SysData.DataColumn avgFCColumn1 = new SysData.DataColumn("AverageFC1", Type.GetType("System.Double"));
            SysData.DataColumn madFCColumn1 = new SysData.DataColumn("MadFC1", Type.GetType("System.Double"));
            SysData.DataColumn avgPColumn1 = new SysData.DataColumn("AverageP1", Type.GetType("System.Double"));


            SysData.DataColumn statColumn2 = new SysData.DataColumn("Mode2", Type.GetType("System.String"));
            SysData.DataColumn cntColumn2 = new SysData.DataColumn("Count2", Type.GetType("System.Int16"));
            SysData.DataColumn percColumn2 = new SysData.DataColumn("Perc2", Type.GetType("System.Double"));
            SysData.DataColumn avgFCColumn2 = new SysData.DataColumn("AverageFC2", Type.GetType("System.Double"));
            SysData.DataColumn madFCColumn2 = new SysData.DataColumn("MadFC2", Type.GetType("System.Double"));
            SysData.DataColumn avgPColumn2 = new SysData.DataColumn("AverageP2", Type.GetType("System.Double"));


            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(cntColumnA);

            lTable.Columns.Add(statColumn1);
            lTable.Columns.Add(cntColumn1);
            lTable.Columns.Add(percColumn1);
            lTable.Columns.Add(avgFCColumn1);
            lTable.Columns.Add(madFCColumn1);
            lTable.Columns.Add(avgPColumn1);

            lTable.Columns.Add(statColumn2);
            lTable.Columns.Add(cntColumn2);
            lTable.Columns.Add(percColumn2);
            lTable.Columns.Add(avgFCColumn2);
            lTable.Columns.Add(madFCColumn2);
            lTable.Columns.Add(avgPColumn2);


            for (int i=0;i<_tmp.Count;i++)
            {
                bool swapped = false;
                SysData.DataRow lRow = lTable.Rows.Add();
                string catName = _tmp[i].catName;
                int totnrgenes = _tmp[i].genes.Length;
                summaryInfo _pos = element_info.Pos.GetCatValues(catName);
                summaryInfo _neg = element_info.Neg.GetCatValues(catName);
                summaryInfo _si1 = _pos;
                summaryInfo _si2 = _neg;
                
                lRow["Name"] = StripText(catName);
                lRow["CountT"] = totnrgenes;
                
                if(_pos.genes.Length < _neg.genes.Length)
                {
                    _si1 = _neg;
                    _si2 = _pos;
                    swapped = true;
                    
                }

                int n1 = _si1.genes.Length;
                //int n2 = _si2.genes[0] == "" ? 0 : _si2.genes.Length;
                int n2 = _si2.genes.Length;

                if(n1==n2) // check for highest FC
                {
                    if(Math.Abs(_si2.fc_average) > Math.Abs(_si1.fc_average))
                    {
                        _si1 = _neg;
                        _si2 = _pos;
                        swapped = !swapped;
                    }
                }

                lRow["Mode1"] = swapped ? "repressed" : "activated";
                lRow["Count1"] = _si1.genes.Length;
                lRow["Perc1"] = totnrgenes == 0 ? 0:Math.Round(((double)n1 / (double)totnrgenes) * 100, 0);
                if (n1 > 0)
                {
                    lRow["AverageFC1"] = _si1.fc_average;
                    lRow["MadFC1"] = _si1.fc_mad;
                    lRow["AverageP1"] = _si1.p_average;
                }

                lRow["Mode2"] = swapped ? "activated" : "repressed";
                lRow["Count2"] = _si2.genes.Length;
                lRow["Perc2"] = totnrgenes==0?0:Math.Round(((double)n2 / (double)totnrgenes) * 100, 0);
                if (n2 > 0)
                {
                    lRow["AverageFC2"] = _si2.fc_average;
                    lRow["MadFC2"] = _si2.fc_mad;
                    lRow["AverageP2"] = _si2.p_average;
                }
            }
            
            return lTable;
        }


        private DataTable ElementsToTable(List<summaryInfo> elements)
        {

            SysData.DataTable lTable = new SysData.DataTable("Elements");
            
            SysData.DataColumn regColumn = new SysData.DataColumn("Name", Type.GetType("System.String"));
            SysData.DataColumn cntColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
            SysData.DataColumn madColumn = new SysData.DataColumn("Mad", Type.GetType("System.Double"));


            //SysData.DataColumn stdColumn = new SysData.DataColumn("Std", Type.GetType("System.Double"));
            SysData.DataColumn avgPColumn = new SysData.DataColumn("P_Average", Type.GetType("System.Double"));
            //SysData.DataColumn madPColumn = new SysData.DataColumn("P_Mad", Type.GetType("System.Double"));


            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(cntColumn);
            lTable.Columns.Add(avgColumn);
            lTable.Columns.Add(madColumn);
            //lTable.Columns.Add(stdColumn);
            lTable.Columns.Add(avgPColumn);
            //lTable.Columns.Add(madPColumn);


            for (int r = 0; r < elements.Count; r++)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                string name = elements[r].catName;
                string[] names = name.Split('(');
                int hit = names[0].ToUpper().IndexOf("REGULON");
                string newname = hit == -1 ? names[0] : names[0].Substring(0, hit);

                //if (mode == ELTYPE.ALL)
                {

                    lRow["Name"] = newname; // name.Substring(0, names[0].Length);
                    lRow["Count"] = elements[r].p_values == null ? 0 : elements[r].p_values.Count();
                    if (!(elements[r].fc_average is Double.NaN))
                        lRow["Average"] = elements[r].fc_average;
                    if(!(elements[r].fc_mad is Double.NaN))
                        lRow["Mad"] = elements[r].fc_mad.ToString();
                    if (!(elements[r].p_average is Double.NaN))
                        lRow["P_Average"] = elements[r].p_average;
                }             
            }

            return lTable;

        }


        private void CreateExtendedRegulonCategoryDataSheet(element_fc theElements,string chartName)
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

        private Excel.Worksheet CreateRankingDataSheet(element_fc theElements, List<summaryInfo> all, List<summaryInfo> posSort, List<summaryInfo> negSort, List<summaryInfo> comSort) 
        {
            string catRegLabel = Properties.Settings.Default.useCat ? "CatRankTab_" : "RegRankTab_";
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, catRegLabel);

            DataTable lTable = ElementsToTable(all);

            string catRegHeader = Properties.Settings.Default.useCat ? "Category" : "Regulon";

            int hdrRow = 2;


            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[1, 5];
            Excel.Range rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
            rall.Merge();
            rall.Value = "PLOT DATA";
            rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            lNewSheet.Cells[hdrRow, 1] = catRegHeader;
            lNewSheet.Cells[hdrRow, 2] = "Nr Genes";
            lNewSheet.Cells[hdrRow, 3] = "Average FC";
            lNewSheet.Cells[hdrRow, 4] = "MAD FC";
            //lNewSheet.Cells[1, 5] = "STD FC";
            lNewSheet.Cells[hdrRow, 5] = "Average P";
            //lNewSheet.Cells[1, 7] = "MAD P";

            // Sort the data with ascending p-values
            DataView lView = lTable.DefaultView;
            //lView.Sort = "P_Average asc";

            // starting from row 2
            FastDtToExcel(lView.ToTable(), lNewSheet, hdrRow+1, 1, lTable.Rows.Count + hdrRow, lTable.Columns.Count);

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
            rall.Value = "POSITIVE FC";
            rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            lTable = ElementsToTable(posSort);

            lNewSheet.Cells[hdrRow, 7] = catRegHeader;
            lNewSheet.Cells[hdrRow, 8] = "Nr Genes Positive";
            lNewSheet.Cells[hdrRow, 9] = "Average FC Positive";
            lNewSheet.Cells[hdrRow, 10] = "MAD FC Positive";            
            lNewSheet.Cells[hdrRow, 11] = "Average P Positive";            

            // Sort the data with ascending p-values
            lView = lTable.DefaultView;
            //lView.Sort = "P_Average asc";

            FastDtToExcel(lView.ToTable(), lNewSheet, hdrRow + 1, 7, lTable.Rows.Count + hdrRow, lTable.Columns.Count+6);

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
            rall.Value = "NEGATIVE FC";
            rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            lNewSheet.Cells[hdrRow, 13] = catRegHeader;
            lNewSheet.Cells[hdrRow, 14] = "Nr Genes Negative";
            lNewSheet.Cells[hdrRow, 15] = "Average FC Negative";
            lNewSheet.Cells[hdrRow, 16] = "MAD FC Negative";
            lNewSheet.Cells[hdrRow, 17] = "Average P Negative";

            // Sort the data with ascending p-values
            lView = lTable.DefaultView;
            //lView.Sort = "P_Average asc";

            FastDtToExcel(lView.ToTable(), lNewSheet, hdrRow + 1, 13, lTable.Rows.Count + hdrRow, lTable.Columns.Count + 12);

            top = lNewSheet.Cells[1, 13];
            bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 17];
            rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
            rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4;
            rall.Interior.TintAndShade = 0.8;
            rall.Interior.PatternTintAndShade = 0;



            top = lNewSheet.Cells[1, 19];
            bottom = lNewSheet.Cells[1, 32];
            rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
            rall.Merge();
            rall.Value = "COMBINED RESULTS";

            rall.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            lNewSheet.Cells[hdrRow, 19] = catRegHeader;
            lNewSheet.Cells[hdrRow, 20] = "Total nr genes";
            lNewSheet.Cells[hdrRow, 21] = "Status";
            lNewSheet.Cells[hdrRow, 22] = "Nr of genes";
            lNewSheet.Cells[hdrRow, 23] = "%";
            lNewSheet.Cells[hdrRow, 24] = "Average FC";
            lNewSheet.Cells[hdrRow, 25] = "MAD FC";
            lNewSheet.Cells[hdrRow, 26] = "Average P";
            lNewSheet.Cells[hdrRow, 27] = "Status";
            lNewSheet.Cells[hdrRow, 28] = "Nr of genes";
            lNewSheet.Cells[hdrRow, 29] = "%";
            lNewSheet.Cells[hdrRow, 30] = "Average FC";
            lNewSheet.Cells[hdrRow, 31] = "MAD FC";
            lNewSheet.Cells[hdrRow, 32] = "Average P";


            lTable = UDElementsToTable(theElements);
            


            FastDtToExcel(lTable, lNewSheet, hdrRow + 1, 19, lTable.Rows.Count + hdrRow, lTable.Columns.Count + 18);

            top = lNewSheet.Cells[1, 19];
            bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 32];
            rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
            rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
            rall.Interior.TintAndShade = 0.8;
            rall.Interior.PatternTintAndShade = 0;





            return lNewSheet;

        }


        private void CreateQPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "QPlot");

#region format_data
            SysData.DataTable _fc_BSU = ReformatResults(aOutput);
            HashSet<string> lRegulons = new HashSet<string>();

            SysData.DataView lRelevant = aSummary.AsDataView();
            lRelevant.RowFilter = "totrel>0";
            SysData.DataTable dataTable = lRelevant.ToTable();


            foreach (SysData.DataRow row in dataTable.Rows)
            {
                lRegulons.Add(row.ItemArray[0].ToString());
            }

            string subsets = string.Join(",", lRegulons.ToArray());
            subsets = string.Join(",", subsets.Split(',').Select(x => $"'{x}'"));

            SysData.DataView dataView = _fc_BSU.AsDataView();
            dataView.RowFilter = String.Format("Regulon in ({0})", subsets);
            dataTable = dataView.ToTable();
#endregion


            //Excel.Shape qPlot = enrichmentAnalysis1.DrawQPlot(lRegulons, dataTable);

            //qPlot.Name = "qPlot";
            //qPlot.Copy();

            //Excel.Range aRange = lNewSheet.Cells[1, 1];
            //lNewSheet.Paste(aRange);

            //foreach (Excel.Shape aShape in lNewSheet.Shapes)
            //{
            //    if (aShape.Name == "qPlot")
            //    {
            //        aShape.Top = 10;
            //        aShape.Left = 10;
            //        aShape.Width = 900;
            //        aShape.Height = 800;

            //    }
            //}

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }

#endregion

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //if(gLastFolder!="")
                
                openFileDialog.InitialDirectory = gLastFolder;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.referenceFile = openFileDialog.FileName;
                    btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;
                    load_Worksheets();
                    btLoad.Enabled = true;

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.referenceFile);
                    gLastFolder = fInfo.DirectoryName;
                }
            }
        }

        private void btnSelectOperonFile_Click(object sender, RibbonControlEventArgs e)
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
                    load_OperonSheet();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.operonFile);
                    gLastFolder = fInfo.DirectoryName;

                }
            }
        }

        private void ddGene_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceGene = ddGene.SelectedItem.Label;
            SetFlags(UPDATE_FLAGS.ALL);
        }
       
        private void editMinPval_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (double.TryParse(editMinPval.Text, out double val))
            {
                // set the text value to what is parsed
                editMinPval.Text = val.ToString();                
                Properties.Settings.Default.pvalue_cutoff = val;
                
                SetFlags(UPDATE_FLAGS.P_dependent);
                //gpValueUpdate = true;
            }
            else            
                editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();                            
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            //splitButton3.Label = but_pvalues.Label;
            //splitButton3.Image = but_pvalues.Image;
            Properties.Settings.Default.use_pvalues = true;
            SetFlags(UPDATE_FLAGS.P_dependent);
        }

        private void but_fc_Click(object sender, RibbonControlEventArgs e)
        {
            //splitButton3.Label = but_fc.Label;
            //splitButton3.Image = but_fc.Image;
            Properties.Settings.Default.use_pvalues = false;
            SetFlags(UPDATE_FLAGS.FC_dependent);
        }

        private void tglTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            var taskpane = TaskPaneManager.GetTaskPane("A", "GIN tool manual", () => new GINtaskpane(), SetTaskPaneVisbile);
            taskpane.Visible = !taskpane.Visible;            
        }


        public void SetTaskPaneVisbile(bool visible)
        {
            tglTaskPane.Checked = visible;
        }

        private void btnResetOperonFile_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.operonFile = "";
            Properties.Settings.Default.operonSheet = "";
            btnOperonFile.Label = "No file selected";
           
            gOperonOutput = false;
            cbOperon.Checked = false;
            cbOperon.Enabled = false;
            Properties.Settings.Default.tblOperon = false;
        }

        private void btnEA_Click(object sender, RibbonControlEventArgs e)
        {
            

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "Plots_");
            //PlotRoutines enrichmentAnalysis = new PlotRoutines(gApplication);
            
            //Excel.Shape distPlot = enrichmentAnalysis.DrawEnrichmentChart();
            //distPlot.Name = "distributionPlot";
            //distPlot.Copy();
            
            //Excel.Range aRange = lNewSheet.Cells[4, 4];
            //lNewSheet.Paste(aRange);

            //Excel.Shape dc2 = distPlot.Duplicate();
            //dc2.Name = "otherPlot";
            //dc2.Copy();
            //lNewSheet.Paste();


            //float p0_height = 0;
            //float p0_width = 0;


            //int shapenr = 0;
            //foreach(Excel.Shape aShape in lNewSheet.Shapes)
            //{
               
            //    if (aShape.Name == "distributionPlot")
            //    {
            //        aShape.Top = 10;
            //        aShape.Left = 100;
            //        aShape.Width = 500;
            //        aShape.Height = 300;

            //        p0_height = aShape.Height;
            //        p0_width = aShape.Width;
            //    }
            //    else
            //    {
            //        aShape.Top = 10;
            //        aShape.Left = p0_width+100;
            //        aShape.Height = p0_height;
            //    }

            //    shapenr++;

            //}


            //dc2.Top = 10;
            //dc2.Left = 600; // distPlot.Width+distPlot.Left;
            //dc2.Width = 100;


        }

        //private void clrExcel_Click(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrExcel.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Excel;
        //}

        //private void clrGray_Click(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrGray.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Grayscale;

        //}


        private void cbGeneratePlots(object sender, RibbonControlEventArgs e)
        {
            //gGeneratePlots = cbGenPlots.Checked;
            //Properties.Settings.Default.generatePlots = gGeneratePlots;
            //if (gGeneratePlots)
            //    if (gEnrichmentAnalysis == null)
            //    {
            //        gEnrichmentAnalysis = new EnrichmentAnalysis(gApplication);
            //    }
        }

        private void cbOrderFC_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.useSort= cbOrderFC.Checked;
            //gOrderAscending = cbOrderFC.Checked;
        }

        private void btnSelectCatFile_Click(object sender, RibbonControlEventArgs e)
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
                    load_CatFile();

                    System.IO.FileInfo fInfo = new System.IO.FileInfo(Properties.Settings.Default.categoryFile);
                    gLastFolder = fInfo.DirectoryName;

                }
            }
        }

        private void ddCatLevel_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            //gDDcatLevel = ddCatLevel.SelectedItemIndex + 1;
        }

        private void cbUseCategories_Click(object sender, RibbonControlEventArgs e)
        {
            //Properties.Settings.Default.catPlot = cbUseCategories.Checked;
            Properties.Settings.Default.useCat = cbUseCategories.Checked;
            cbUseRegulons.Checked = !cbUseCategories.Checked;

        }

        private void cbClustered_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.catPlot = cbClustered.Checked;
        }

        private void cbDistribution_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.distPlot = cbDistribution.Checked;
        }

        // data for regulon plot
        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.regPlot = chkRegulon.Checked;            
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            ShowSettingPannels(toggleButton1.Checked);
            grpPlot.Visible = !toggleButton1.Checked;
            grpTable.Visible = !toggleButton1.Checked;
            grpDta.Visible = !toggleButton1.Checked;
        }

        private void ShowSettingPannels(bool show)
        {
            grpReference.Visible = show;
            grpMap.Visible = show;
            grpUpDown.Visible = show;
            grpFC.Visible = show;
            grpCutOff.Visible = show;
            grpDirection.Visible = show;
            //grpTables.Visible = show;

        }

        


        //private void cbQplot_Click(object sender, RibbonControlEventArgs e)
        //{
        //    gQPlot = cbQplot.Checked;
        //    Properties.Settings.Default.qPlot = gQPlot;

        //}

        //private void clrGray_Click_1(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrGray.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Grayscale;
        //}

        //private void clrBerry_Click(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrBerry.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Berry;
        //}

        //private void clrBright_Click(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrBright.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Bright;
        //}

        //private void clrBrightPastel_Click(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrBrightPastel.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.BrightPastel;
        //}

        //private void clrChocolate_Click(object sender, RibbonControlEventArgs e)
        //{
        //    btnPalette.Image = clrChocolate.Image;
        //    Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Chocolate;
        //}

        public enum TASKS : int
        {
            READY = 0,
            LOAD_REGULON_DATA,
            LOAD_OPERON_DATA,
            LOAD_CATEGORY_DATA,
            MAPPING_GENES_TO_REGULONS,
            READ_SHEET_DATA,
            READ_SHEET_CAT_DATA,
            UPDATE_MAPPED_TABLE,
            UPDATE_SUMMARY_TABLE,
            UDPATE_COMBINED_TABLE,
            UPDATE_OPERON_TABLE,  // a table for now.. should become table & graph           
            COLOR_CELLS,
            CATEGORY_CHART,
            REGULON_CHART
        };

        public string[] taks_strings = new string[] { "Ready", "Load regulon data", "Load operon data", "Load category data", "Mapping genes to regulons",  "Read sheet data", "Read sheet categorized data", "Update mapping table", "Update summary table", "Update combined table", "Update operon table" , "Color cells", "Create category chart", "Create regulon chart" };


        public enum UPDATE_FLAGS:byte
        {     
            TSummary  = 0b_0000_0001,
            TCombined = 0b_0000_0010,
            TOperon   = 0b_0000_0100,
            TMapped   = 0b_0000_1000,
            PRegulon  = 0b_0001_0000,
            PDist     = 0b_0010_0000,
            PCat      = 0b_0100_0000,
            POperon   = 0b_1000_0000,
            
            FC_dependent = TCombined | POperon | TSummary, 
            P_dependent = TCombined | POperon | TSummary,
            
            ALL = 0b_1111_1111,
            NONE = 0b_0000_0000
        };


        public enum FLAG_BITS : int
        {
            TSummary = 0,
            TCombined = 1,
            TOperon = 2,
            TMapped = 3,
            PRegulon = 4,
            PDist = 5,
            PCat = 6,
            POperon = 7,                        
        };

        private enum SORTMODE : int
        {
            FC = 0,
            P = 1,
            MADP = 2,
            NGENES = 3,
            CATNAME = 4,
        };

        private void cbMapping_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblMap = cbMapping.Checked;
        }

        private void cbSummary_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblSummary = cbSummary.Checked;
        }

        private void cbCombined_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblCombine = cbCombined.Checked;
        }

        private void cbOperon_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.tblOperon = cbOperon.Checked;
        }

        private void btnClearCatFile_Click(object sender, RibbonControlEventArgs e)
        {
            gCatOutput = false;
            cbUseCategories.Checked = false;
            cbUseCategories.Enabled = false;           
            Properties.Settings.Default.useCat = false;
            cbUseRegulons.Checked = true;

            Properties.Settings.Default.categoryFile= "";            
            btnCatFile.Label = "No file selected";



        }

        private void cbUsePValues_Click(object sender, RibbonControlEventArgs e)
        {
            cbUseFoldChanges.Checked = !cbUsePValues.Checked;
            Properties.Settings.Default.use_pvalues = cbUsePValues.Checked;
        }

        private void cbUseFoldChanges_Click(object sender, RibbonControlEventArgs e)
        {
            cbUsePValues.Checked = !cbUseFoldChanges.Checked;
            Properties.Settings.Default.use_pvalues = cbUsePValues.Checked;        
        }

        private void btnSelect_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Range theInputCells = GetActiveCell();

            dlgSelectData sd = new dlgSelectData(theInputCells);
            sd.theApp = gApplication;
            if(sd.ShowDialog() == DialogResult.OK)
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
            //gApplication.EnableEvents = false;
            //gApplication.DisplayAlerts = false;


            //gApplication.EnableEvents = true;
            //gApplication.DisplayAlerts = true;

        }

        private void cbDescending_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.sortAscending = !cbDescending.Checked;            
            cbAscending.Checked = !cbDescending.Checked;
        }

        private void cbAscending_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.sortAscending = cbAscending.Checked;
            cbDescending.Checked = !cbAscending.Checked;
        }

        private void cbUseRegulons_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.useCat = !cbUseRegulons.Checked;
            cbUseCategories.Checked = !cbUseRegulons.Checked;
        }
    }


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
