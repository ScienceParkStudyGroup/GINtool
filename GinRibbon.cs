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

        bool gOperonOutput = false;
        bool gCatOutput = false;
        //bool gpValueUpdate = true;
        //bool gRegulonPlot = true;

        List<chart_info> gCharts = new List<chart_info>();

        byte gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;

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

        string gInputRange = "";
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

            gApplication.StatusBar = "Load category data";

            gApplication.EnableEvents = false;

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
            gApplication.EnableEvents = true;
            gApplication.StatusBar = "Ready";            

            return gCategories.Rows.Count > 0;
        }

        private bool LoadOperonData()
        {
          
            if (Properties.Settings.Default.operonFile.Length == 0 || Properties.Settings.Default.operonSheet.Length == 0)            
                return false;            

            gApplication.EnableEvents = false;
            gApplication.StatusBar = "Load operon data";

            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.operonSheet, Properties.Settings.Default.operonFile, 1, 1);
            gRefOperons = new SysData.DataTable("OPERONS");
            gRefOperons.CaseSensitive = false;
            gRefOperons.Columns.Add("operon", Type.GetType("System.String"));
            gRefOperons.Columns.Add("gene", Type.GetType("System.String"));

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
                }
            }
            gApplication.EnableEvents = true;
            gApplication.StatusBar = "Ready";
            return gRefOperons.Rows.Count>0;
        }

        private bool LoadData()
        {
            gApplication.StatusBar = "Load regulon/gene mappings";
            gApplication.EnableEvents = false;
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
            gApplication.EnableEvents = true;
            gApplication.StatusBar = "Ready";
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

        private void GinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            gApplication = Globals.ThisAddIn.GetExcelApplication();
            btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;
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
            chkRegulon.Checked = Properties.Settings.Default.regPlot;

            //cbOrderFC.Checked = Properties.Settings.Default.useFCorder;
            cbUseCategories.Checked = Properties.Settings.Default.useCat;
            cbUsePValues.Checked = Properties.Settings.Default.use_pvalues;
            cbUseFoldChanges.Checked = !Properties.Settings.Default.use_pvalues;
            

            if (Properties.Settings.Default.operonFile.Length == 0)
                btnOperonFile.Label = "No file selected";

            gAvailItems = propertyItems("directionMapUnassigned");
            gUpItems = propertyItems("directionMapUp");
            gDownItems = propertyItems("directionMapDown");

            btnSelect.Enabled = false;
            btApply.Enabled = false;
            ddBSU.Enabled = false;
            ddGene.Enabled = false;
            ddRegulon.Enabled = false;
            ddDir.Enabled = false;
            btPlot.Enabled = false;
            cbUseCategories.Enabled = false;
            cbMapping.Enabled = false;
            cbSummary.Enabled = false;
            cbCombined.Enabled = false;
            cbOperon.Enabled = false;
            cbOrderFC.Enabled = false;
            cbUsePValues.Enabled = false;
            cbUseFoldChanges.Enabled = false;
            toggleButton1.Enabled = true;
            cbAscending.Enabled = false;
            cbDescending.Enabled = false;

            //gOrderAscending = cbOrderFC.Checked;
            //edtMaxGroups.Enabled = false;
            //btnPalette.Enabled = false;

            //cbGenPlots.Checked = Properties.Settings.Default.generatePlots;
            //cbQplot.Checked = Properties.Settings.Default.qPlot;
            //gCompositPlot = cbGenPlots.Checked;
            //gQPlot = cbQplot.Enabled;

            //if (cbGenPlots.Checked )
            //if (gEnrichmentAnalysis == null)
            //{
            //    gEnrichmentAnalysis = new PlotRoutines(gApplication);
            //}

            PlotRoutines.theApp = gApplication;

            EnableOutputOptions(false);

            gExcelErrorValues = ((int[])Enum.GetValues(typeof(ExcelUtils.CVErrEnum))).ToList();

            //if (Properties.Settings.Default.use_pvalues)
            //{
            //    splitButton3.Label = but_pvalues.Label;
            //    splitButton3.Image = but_pvalues.Image;
            //}
            //else
            //{
            //    splitButton3.Label = but_fc.Label;
            //    splitButton3.Image = but_pvalues.Image;
            //}

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
            chkRegulon.Enabled= enable && gOperonOutput;

            cbOrderFC.Enabled = enable;
            cbUseCategories.Enabled = enable && gCatOutput;

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

        private List<BsuRegulons> QueryResultTable(Excel.Range theCells)
        {
            gApplication.StatusBar = "Mapping genes to regulons";
            List<BsuRegulons> lList = new List<BsuRegulons>();

            foreach (Excel.Range c in theCells.Rows)
            {
                string lBSU;                
                double lFC = 0;
                double lPvalue = 1;
                BsuRegulons lMap = null;
               
                if (c.Columns.Count == 3)
                {
                    object[,] value = c.Value2;
                    
                    // first check if the cell contains an erroneous value, if not then try to parse the value or reset to default

                    if (!iserrorCell(value[1, 1]))
                        if (!Double.TryParse(value[1, 1].ToString(), out lPvalue))
                            lPvalue = 1;

                    if (!iserrorCell(value[1, 2]))
                        if (!Double.TryParse(value[1, 2].ToString(), out lFC))
                            lFC = 0;
                                        
                    lBSU = value[1, 3].ToString();
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

            gApplication.StatusBar = "Ready";
            return lList;

        }
     
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



        private void AddTask()
        {

        }

        private void RemoveTask()
        {

        }


        // the main routine after mouse selection update // generates mapping output.. should be de-coupled (update data & mapping output)
        private (List<FC_BSU>, List<BsuRegulons>) GenerateOutput(bool suppressOutput=false)
        {
            gApplication.StatusBar = "Formatting output results";

            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

            Excel.Range theInputCells = GetActiveCell();


            Excel.Worksheet theSheet = GetActiveSheet();

            if (theSheet == null)
                return (null,null);

            if (theSheet.Name.Contains("Plot_"))
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);             
            }

            if (theSheet.Name.Contains("CongruenceData_"))
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);
            }

            if (theSheet.Name.Contains("CongruencePlot_"))
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);
            }

            if (theSheet.Name.Contains("Summary_"))
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);
            }

            if (theSheet.Name.Contains("Combined_"))
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);
            }


            if (theSheet.Name.Contains("Mapped_"))
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);
            }


            if (RangeAddress(theInputCells) != gInputRange || (gOutput == null || gList == null))
            {
                gInputRange = RangeAddress(theInputCells);
                gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;
                //gNeedsUpdating = true;                
            }
            else
            {
                gApplication.ScreenUpdating = true;
                gApplication.DisplayAlerts = true;
                gApplication.EnableEvents = true;
                gApplication.StatusBar = "Ready";
                return (gOutput, gList);
            }
            
            int nrRows = theInputCells.Rows.Count;
            int startC = theInputCells.Column;
            int startR = theInputCells.Row;

            // from now always assume 3 columns.. p-value, fc, bsu
           // int offsetColumn = 1;

            if(theInputCells.Columns.Count !=3)
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);

            }                       

            // generate the results for outputting the data and summary
            List<BsuRegulons> lResults = QueryResultTable(theInputCells);
            // output the data
            //var lOut = PrepareResultTable(lResults);

            //SysData.DataTable lTable = lOut.Item1;
            //SysData.DataTable clrTbl;

            //if (!suppressOutput)
            //{
            //    gApplication.StatusBar = "Creating mapping sheet";

            //    Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            //    renameWorksheet(lNewSheet, "Mapped_");


            //    lNewSheet.Cells[1, 1] = "BSU";
            //    lNewSheet.Cells[1, 2] = "GENE";
            //    lNewSheet.Cells[1, 3] = "FC";
            //    lNewSheet.Cells[1, 4] = "PVALUE";
            //    lNewSheet.Cells[1, 5] = "TOT REGULONS";

            //    string lastColumn = lTable.Columns[lTable.Columns.Count - 1].ColumnName;
            //    lastColumn = lastColumn.Replace("col_", "");
            //    int maxreg = Int16.Parse(lastColumn);

            //    for (int i = 0; i < maxreg; i++)
            //        lNewSheet.Cells[1, i + 6] = string.Format("Regulon_{0}", i + 1);

            //    FastDtToExcel(lTable, lNewSheet, startR, offsetColumn, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);

            //    Excel.Range top = lNewSheet.Cells[1, 1];
            //    Excel.Range bottom = lNewSheet.Cells[lTable.Rows.Count + 1, lTable.Columns.Count];
            //    Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            //    all.Columns.AutoFit();

            //    clrTbl = lOut.Item2;
            //    ColorCells(clrTbl, lNewSheet, startR, offsetColumn + 5, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);

            //}

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



            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;


            gApplication.StatusBar = "Ready";
            


            return (lOutput, lResults);
           
        }
       

        private void CreateMappingSheet(List<BsuRegulons> bsuRegulons)
        {
            var lOut = PrepareResultTable(bsuRegulons);

            SysData.DataTable lTable = lOut.Item1;
            SysData.DataTable clrTbl;
            

            gApplication.StatusBar = "Creating mapping sheet";

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

            clrTbl = lOut.Item2;
            ColorCells(clrTbl, lNewSheet, startR, offsetColumn + 5, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);
            
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
            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

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

            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;

        }


        private void CreateSummarySheet(SysData.DataTable theTable)
        {
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

        }

        private SysData.DataTable ReformatResults(List<FC_BSU> aList)
        {
            // find unique regulons

            SysData.DataTable lTable = new SysData.DataTable("FC_BSU");
            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn geneColumn = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            SysData.DataColumn pvalColumn = new SysData.DataColumn("Pvalue", Type.GetType("System.Single"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Single"));
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
                float fcGene = (float)aRow[i]["FC"];
                int dirBSU = (int)aRow[i]["DIR"];
                float lowValue = Properties.Settings.Default.fcLOW;

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
            int s = 1;
            while (currentSheets.Contains(string.Format("{0}_{1}", wsBase, s)))
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


        //private Excel.Worksheet CreateCongruenceDataSheet(HashSet<string> aRegulons, SysData.DataTable aTable)
        //{

        //    gApplication.StatusBar = "Generating sheet for plotting data";
        //    gApplication.ScreenUpdating = false;
        //    gApplication.DisplayAlerts = false;
        //    gApplication.EnableEvents = false;

        //    Excel.Worksheet aSheet = gApplication.Worksheets.Add();
            
                        

        //    int nrGenes = aTable.Rows.Count;
        //    int nrRegulons = aRegulons.Count;

        //    List<float[]> fc = new List<float[]>();

        //    SysData.DataView dataView = aTable.AsDataView();

        //    double MMAX = (double)(float)aTable.Rows[0]["FC"];
        //    double MMIN = (double)(float)aTable.Rows[0]["FC"];
        //    List<double> meanFC = new List<double>();

        //    foreach (string regulon in aRegulons)
        //    {
        //        dataView.RowFilter = String.Format("Regulon = '{0}'", regulon);
        //        SysData.DataTable dataTable = dataView.ToTable();
        //        int nrRows = dataTable.Rows.Count;
        //        float[] vs = new float[nrRows];
        //        int[] ys = new int[nrRows];
        //        for (int _r = 0; _r < nrRows; _r++)
        //        {
        //            double _val = (double)(float)dataTable.Rows[_r]["FC"];
        //            if (_val > MMAX) { MMAX = _val; }
        //            if (_val < MMIN) { MMIN = _val; }
        //            vs[_r] = (float)_val;
        //            ys[_r] = fc.Count;

        //        }
        //        meanFC.Add(vs.Average());
        //        fc.Add(vs);
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
        //    for (int i = 0; i < nrRegulons; i++)
        //    {
        //        aSheet.Cells[1, 1] = "yoffset";
        //        aSheet.Cells[1, i + 2] = aRegulons.ToArray()[sortedEntries[i]];

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

            lNewSheet.Cells[1, col++] = "Gene";
            lNewSheet.Cells[1, col++] = "Operon Name";
            lNewSheet.Cells[1, col++] = "Nr genes";
            lNewSheet.Cells[1, col++] = "Operon";

            for (int c=0;c< maxNrGenes; c++)
            {
                string colHeader = string.Format("FC Gene #{0}", c + 1);
                lNewSheet.Cells[1, c+col] = colHeader;
            }

            FastDtToExcel(table, lNewSheet, 2, 1, table.Rows.Count + 1, maxNrGenes + 4);


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


        private (List<string>,List<double>) GetOperonGenesFC(string operon, List<BsuRegulons> lLst)
        {

            SysData.DataRow[] lquery = gRefOperons.Select(string.Format("Operon = '{0}'", operon));

            List<string> lgenes = new List<string>();
            List<double> lFCs = new List<double>();

            int nrHits = lquery.Count();

            foreach (DataRow row in lquery)
            {                
                string lgene = row["gene"].ToString();
//                if (lgene == operon && nrHits>1)
//                    continue;
                lgenes.Add(lgene);
                BsuRegulons result = lLst.Find(item => item.GENE == lgene );
                if (result != null)
                    lFCs.Add(result.FC);
                else
                    lFCs.Add(Double.NaN);

            }   
                                   
            return (lgenes,lFCs);
        }

        private SysData.DataTable CreateOperonTable(SysData.DataTable aUsageTbl, List<BsuRegulons> lLst)
        {
            SysData.DataTable lTable = new SysData.DataTable();

            SysData.DataColumn col = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("operon", Type.GetType("System.String"));
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


            //SysData.DataTable uOperons = GetDistinctRecords(gRefOperons, new string[] { "Operon" });
            double lowVal = Properties.Settings.Default.fcLOW;
            for (int r = 0; r < lLst.Count; r++)
            {
                string geneName = lLst[r].GENE;

                bool accept = Properties.Settings.Default.use_pvalues ? lLst[r].PVALUE < Properties.Settings.Default.pvalue_cutoff : Math.Abs(lLst[r].FC) > lowVal;
                if (!accept)
                    continue;

                HashSet<string> luOperons = new HashSet<string>();

                SysData.DataRow[] lOperons = gRefOperons.Select(string.Format("gene='{0}'", lLst[r].GENE));
                foreach (DataRow row in lOperons)
                {
                    string operon = row["operon"].ToString();
                    luOperons.Add(operon);
                }

                foreach(string operon in luOperons)
                { 
                    
                    (List<string> lgenes, List<double> lFCs) = GetOperonGenesFC(operon, lLst);
                    int nrgenes = lgenes.Count;
                    string opgenes = string.Join("-", lgenes.ToArray());
                    // addrow
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["gene"] = geneName;
                    lRow["operon"] = operon;
                    lRow["nrgenes"] = nrgenes;
                    lRow["operon_genes"] = opgenes;
                    for(int i=0;i<nrgenes;i++)
                    {
                        lRow[string.Format("gene_{0}", i + 1)] = lFCs[i];
                    }
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


                float lFClow = Properties.Settings.Default.fcLOW;
                float lFCmid = Properties.Settings.Default.fcMID;
                float lFChigh = Properties.Settings.Default.fcHIGH;

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
                        float fc = (float)_tmp[_r]["FC"];
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
            gInputRange = null;
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

            if (float.TryParse(bx.Text, out float val))
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
                dlgTreeView dlg = new dlgTreeView();

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
                            CategoryPlot(gOutput, gSummary, dlg.GetSelection());
                            //UnSetFlags(UPDATE_FLAGS.PCat);
                            //gNeedsUpdate.UnSet(UPDATE_FLAGS.PCat);
                        }
                    //}
                    //if ((gOutput != null && gSummary != null && dlg.GetSelection().Count() > 0) ) //&& gNeedsUpdate.Check(UPDATE_FLAGS.PRegulon))
                    //{ 
                        if (Properties.Settings.Default.regPlot)
                        {
                            // data for now.. to be changed in plot
                            RegulonPlotData(gOutput, gSummary, dlg.GetSelection());
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
        //private void btApply_Click(object sender, RibbonControlEventArgs e)
        //{
        //    gApplication.EnableEvents = false;
        //    gApplication.DisplayAlerts = false;

        //    if(gNeedsUpdate.None())            
        //        return;            


        //    if (gOutput == null || gSummary == null)
        //    {

        //        if(gOutput == null || gList ==null || gNeedsUpdating)
        //            (gOutput, gList) = GenerateOutput();

        //        if (gOutput is null || gList is null)
        //            return;

        //        if ((gSummary == null && gCombineInfo == null) || gNeedsUpdating)
        //        {
        //            (gSummary, gCombineInfo) = CreateUsageTable(gOutput);
        //            CreateSummarySheet(gSummary);

        //            //CreateCombinedSheet(lCombined);          
        //        }

        //        if (gpValueUpdate) // add fccheck
        //        {
        //            SysData.DataTable lCombined = CreateCombinedTable(gCombineInfo, gList);
        //            CreateCombinedSheet(lCombined);

        //            if(gOperonOutput)
        //            {
        //                SysData.DataTable tblOperon = CreateOperonTable(gCombineInfo, gList);
        //                CreateOperonSheet(tblOperon);
        //            }    

        //            gpValueUpdate = false;
        //        }

        //        gNeedsUpdating = false;
        //    }

        //    gApplication.EnableEvents = true;
        //    gApplication.DisplayAlerts = true;
        //}
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



        private List<element_fc> CatElements2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements)
        {

            List<element_fc> element_Fcs = new List<element_fc>();
            SysData.DataView dataViewCat = gCategories.AsDataView();
            foreach (cat_elements ce in cat_Elements)
            {
                string categories = string.Join(",", ce.elements.ToArray());
                categories = string.Join(",", categories.Split(',').Select(x => $"'{x}'"));
                dataViewCat.RowFilter = String.Format("catid_short in ({0})", categories);

                HashSet<string> genes = new HashSet<string>();
                foreach(DataRow _row in dataViewCat.ToTable().Rows)
                {
                    genes.Add(_row[gCategoryGeneColumn].ToString());
                }
                
                string genesFormat = string.Join(",", genes.ToArray());
                genesFormat = string.Join(",", genesFormat.Split(',').Select(x => $"'{x}'"));
                dataView.RowFilter = String.Format("Gene in ({0})", genesFormat);

                SysData.DataTable _dt = dataView.ToTable(true, "Gene", "FC");

                element_fc element_Fc;
                element_Fc.catName = string.Format("{0}({1})",ce.catName,_dt.Rows.Count);
                List<float> _fcs = new List<float>();
                foreach(DataRow _row in _dt.Rows)
                {
                    _fcs.Add((float)_row["FC"]);
                }

                if (_fcs.Count == 0)
                {
                    element_Fc.average = 0;
                    element_Fc.fc = null;
                    element_Fc.sd = 0;
                    element_Fc.mad = 0;
                    element_Fc.genes = null;
                }
                else
                {
                    element_Fc.average = _fcs.Average();
                    element_Fc.fc = _fcs.ToArray();
                    element_Fc.sd = _fcs.sd();
                    element_Fc.mad = _fcs.mad();
                    element_Fc.genes = null;
                }
                element_Fcs.Add(element_Fc);                               
            }

            if (Properties.Settings.Default.useSort)
            {
                float[] __values = element_Fcs.Select(x => x.average).ToArray();
                var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<float, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<float, int>(x, i)).OrderByDescending(x => x.Key).ToList();

                //List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                //return (element_Fcs, sortedIndex);

                element_Fcs = sortedElements.Select(x => element_Fcs[x.Value]).ToList();
            }
            else // don't sort but possibly reverse direction
            {
                if(Properties.Settings.Default.sortAscending)
                    element_Fcs.Reverse();

            }

            return element_Fcs;
        }


        private List<element_fc> Regulons2ElementsFC(SysData.DataView dataView, List<cat_elements> cat_Elements) // HashSet<string> regulons)
        {
            List<element_fc> element_Fcs = new List<element_fc>();
            
            foreach(cat_elements el in cat_Elements)
            {                
                dataView.RowFilter = String.Format("Regulon='{0}'", el.catName);
                element_fc element_Fc;
                //element_Fc.catName = el.catName;
                
                SysData.DataTable _dataTable = dataView.ToTable();
                element_Fc.catName = string.Format("{0}({1})", el.catName, _dataTable.Rows.Count);
                if (_dataTable.Rows.Count > 0)
                {
                    List<float> _fcs = new List<float>(_dataTable.Rows.Count);
                    List<string> _genes = new List<string>(_dataTable.Rows.Count);
                    for (int i = 0; i < _dataTable.Rows.Count; i++)
                    {
                        _genes.Add(_dataTable.Rows[i]["Gene"].ToString());
                        _fcs.Add(float.Parse(_dataTable.Rows[i]["FC"].ToString()));
                    }

                    element_Fc.average = _fcs.Average();
                    element_Fc.fc = _fcs.ToArray();
                    element_Fc.sd = _fcs.sd();
                    element_Fc.mad = _fcs.mad();
                    element_Fc.genes = _genes.ToArray();
                    element_Fcs.Add(element_Fc);
                }
                else
                {
                    element_Fc.average = 0;
                    element_Fc.fc = new float[] { 0 };
                    element_Fc.genes = new string[] { "" };
                    element_Fc.sd = 0;
                    element_Fc.mad = 0;
                    element_Fcs.Add(element_Fc);
                }
            }

            if (Properties.Settings.Default.useSort)
            {
                float[] __values = element_Fcs.Select(x => x.average).ToArray();
                var sortedElements = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<float, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<float, int>(x, i)).OrderByDescending(x => x.Key).ToList();

                List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();
                //return (element_Fcs, sortedIndex);
                element_Fcs = sortedElements.Select(x => element_Fcs[x.Value]).ToList();
            }
            else // don't sort but possibly reverse direction
            {
                if (Properties.Settings.Default.sortAscending)
                    element_Fcs.Reverse();

            }
            return element_Fcs;
        }

    
        // output of all genes in table
        private (List<float>, List<int>) SortedFoldChanges(SysData.DataTable dataTable)
        {
            List<float> _values = new List<float>();
            foreach (SysData.DataRow row in dataTable.Rows)
            {
                _values.Add(row.Field<float>("FC"));
            }

            float[] __values = _values.ToArray();
            var sortedGenes = (!Properties.Settings.Default.sortAscending) ? __values.Select((x, i) => new KeyValuePair<float, int>(x, i)).OrderBy(x => x.Key).ToList() : __values.Select((x, i) => new KeyValuePair<float, int>(x, i)).OrderByDescending(x => x.Key).ToList();

            
            List<float> sortedGenesValues = sortedGenes.Select(x => x.Key).ToList();
            List<int> sortedGenesInt = sortedGenes.Select(x => x.Value).ToList();
            return (sortedGenesValues, sortedGenesInt);
        }


        public Excel.Chart MyTmpPlot(List<element_fc> element_Fcs)
        {
            

            if (gApplication == null)
                return null;

            Excel.Worksheet aSheet = gApplication.Worksheets.Add();

            
            //var missing = System.Type.Missing;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
            Excel.Chart chartPage = myChart.Chart;            

            chartPage.ChartType = Excel.XlChartType.xlXYScatter;

            var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

            int nrCategories = element_Fcs.Count;

            float MMAX = 0;
            float MMIN = 0;

            for (int _i = 0; _i < nrCategories; _i++)
            {
                if (element_Fcs[_i].fc != null)
                {
                    if (element_Fcs[_i].fc.Min() < MMIN)
                        MMIN = element_Fcs[_i].fc.Min();
                    if (element_Fcs[_i].fc.Max() > MMAX)
                        MMAX = element_Fcs[_i].fc.Max();
                }
            }


            foreach (var element_Fc in element_Fcs.Select((value, index) => new { value, index }))
            {
                var xy1 = series.NewSeries();
                xy1.Name = element_Fc.value.catName;
                xy1.ChartType = Excel.XlChartType.xlXYScatter;
                if (element_Fc.value.fc != null)
                {
                    xy1.XValues = element_Fc.value.fc;
                    xy1.Values = Enumerable.Repeat(element_Fc.index + 0.5, element_Fc.value.fc.Length).ToArray();
                    xy1.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
                    xy1.MarkerSize = 2;
                    xy1.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, 0.1);
                    Excel.ErrorBars errorBars = xy1.ErrorBars;
                    errorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap;
                    errorBars.Format.Line.Weight = 1.25f;

                    // give each serie different color
                    switch (element_Fc.index % 6)
                    {
                        case 0:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent1;
                            break;
                        case 1:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent2;
                            break;
                        case 2:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent3;
                            break;
                        case 3:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent4;
                            break;
                        case 4:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent5;
                            break;
                        case 5:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = Microsoft.Office.Core.MsoThemeColorIndex.msoThemeColorAccent6;
                            break;
                    }


                }
                var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                //yAxis.AxisTitle.Text = "Regulon";
                Excel.TickLabels labels = yAxis.TickLabels;
                labels.Offset = 1;
            }




            chartPage.ChartColor = 1; // Properties.Settings.Default.defaultPalette;

            // as a last step, add the axis labels series

            if (true)
            {

                var xy2 = series.NewSeries();

                xy2.ChartType = Excel.XlChartType.xlXYScatter;
                //# Excel.Range rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i*2)+1], _tmpSheet.Cells[6, (i * 2) + 1]];
                xy2.XValues = Enumerable.Repeat(MMIN, nrCategories).ToArray();

                //rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i * 2) + 2], _tmpSheet.Cells[6, (i * 2) + 2]];
                float[] yv = new float[nrCategories];
                for (int _i = 0; _i < nrCategories; _i++)
                {
                    yv[_i] = ((float)_i) + 0.5f;
                }

                xy2.Values = yv;

                xy2.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
                xy2.HasDataLabels = true;

                for (int _i = 0; _i < nrCategories; _i++)
                {
                    xy2.DataLabels(_i + 1).Text = element_Fcs[_i].catName;
                }

                xy2.DataLabels().Position = Excel.XlDataLabelPosition.xlLabelPositionLeft;

            }


            chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
            chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();            
            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
            chartPage.Axes(Excel.XlAxisType.xlValue).MaximumScale = nrCategories;
            chartPage.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
            
            chartPage.Legend.Delete();
            chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);
          
            aSheet.Delete();
            return chartPage;

        }


        private void DistributionPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            SysData.DataTable _fc_BSU_ = ReformatResults(aOutput);
            SysData.DataTable _fc_BSU = GetDistinctRecords(_fc_BSU_, new string[] { "Gene","FC"});

            (List<float> sFC, List<int> sIdx) = SortedFoldChanges(_fc_BSU);


            int chartNr = nextWorksheet("DistributionPlot");
            string chartName = "DistributionPlot_" + chartNr.ToString();

            Excel.Chart aChart = PlotRoutines.CreateDistributionPlot(sFC,sIdx, chartName);
            this.RibbonUI.ActivateTab("TabGINtool");


            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }


        private void CategoryPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements )
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;
                       
            SysData.DataTable _fc_BSU = ReformatResults(aOutput);


            List<element_fc> element_Fcs = new List<element_fc>();

            // HashSet ensures unique list
            HashSet<string> lRegulons = new HashSet<string>();

            foreach (SysData.DataRow row in aSummary.Rows)
                lRegulons.Add(row.ItemArray[0].ToString());
            
            SysData.DataView dataView = _fc_BSU.AsDataView();
            List<element_fc> catPlotData = null;            
            if (Properties.Settings.Default.useCat)
            {
                catPlotData = CatElements2ElementsFC(dataView, cat_Elements);                
            }
            else
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements);
           

            int chartNr = Properties.Settings.Default.useCat ? nextWorksheet("CategoryPlot"): nextWorksheet("RegulonPlot");            
            string chartName = (Properties.Settings.Default.useCat ? "CategoryPlot_" : "RegulonPlot_") + chartNr.ToString();


#if CLICK_CHART
            PlotRoutines.CreateCategoryPlot(catPlotData,chartName);
            Excel.Chart aChart = gApplication.ActiveChart;
            aChart.MouseDown += new Excel.ChartEvents_MouseDownEventHandler(AChart_MouseDown);
            gCharts.Add(new chart_info(aChart, catPlotData));
#endif

            this.RibbonUI.ActivateTab("TabGINtool");
            

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;            
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

        private void RegulonPlotData(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            SysData.DataTable _fc_BSU = ReformatResults(aOutput);


            List<element_fc> element_Fcs = new List<element_fc>();

            // HashSet ensures unique list
            HashSet<string> lRegulons = new HashSet<string>();

            foreach (SysData.DataRow row in aSummary.Rows)
                lRegulons.Add(row.ItemArray[0].ToString());

            SysData.DataView dataView = _fc_BSU.AsDataView();
            List<element_fc> catPlotData = null;
            if (Properties.Settings.Default.useCat)
            {
                catPlotData = CatElements2ElementsFC(dataView, cat_Elements);
            }
            else
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements);

            CreateRegulonPlotDataSheet(catPlotData);

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }


        private DataTable ElementsToTable(List<element_fc> elements)
        {

            SysData.DataTable lTable = new SysData.DataTable("Elements");
            SysData.DataColumn regColumn = new SysData.DataColumn("Name", Type.GetType("System.String"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Single"));
            SysData.DataColumn madColumn = new SysData.DataColumn("Mad", Type.GetType("System.Single"));
            SysData.DataColumn stdColumn = new SysData.DataColumn("Std", Type.GetType("System.Single"));
            


            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(avgColumn);
            lTable.Columns.Add(madColumn);
            lTable.Columns.Add(stdColumn);
            

            for (int r = 0; r < elements.Count; r++)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                lRow["Name"] = elements[r].catName;
                lRow["Average"] = elements[r].average;
                lRow["Mad"] = elements[r].mad;
                lRow["Std"] = elements[r].sd;
                
            }

            return lTable;

        }


        private void CreateRegulonPlotDataSheet(List<element_fc> theElements)
        {
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "RegulonImp_");

            DataTable lTable = ElementsToTable(theElements);


            lNewSheet.Cells[1, 1] = "Regulon/Category";
            lNewSheet.Cells[1, 2] = "Average FC";
            lNewSheet.Cells[1, 3] = "MAD FC";

            lNewSheet.Cells[1, 4] = "STD FC";
            
            // starting from row 2


            FastDtToExcel(lTable, lNewSheet, 2, 1, lTable.Rows.Count + 1, lTable.Columns.Count);

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
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.referenceFile = openFileDialog.FileName;
                    btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;
                    load_Worksheets();
                    btLoad.Enabled = true;
                }
            }
        }

        private void btnSelectOperonFile_Click(object sender, RibbonControlEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.operonFile = openFileDialog.FileName;
                    btnOperonFile.Label = Properties.Settings.Default.operonFile;
                    load_OperonSheet();
                    
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
            if (float.TryParse(editMinPval.Text, out float val))
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
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|txt files (*.csv)|*.csv";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.categoryFile = openFileDialog.FileName;
                    btnCatFile.Label = Properties.Settings.Default.categoryFile;
                    load_CatFile();

                }
            }
        }

        private void ddCatLevel_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            //gDDcatLevel = ddCatLevel.SelectedItemIndex + 1;
        }

        private void cbUseCategories_Click(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.useCat = cbUseCategories.Checked;
            
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
            
            FC_dependent = TCombined | POperon,
            P_dependent = TCombined | POperon,
            
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
            //gApplication.EnableEvents = false;
            //gApplication.DisplayAlerts = false;

            (gOutput, gList) = GenerateOutput();

            if (gOutput is null || gList is null)
                return;

            btApply.Enabled = true;
            btPlot.Enabled = true;                 
            EnableOutputOptions(true);
            
            UnSetFlags(UPDATE_FLAGS.TMapped);            

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
