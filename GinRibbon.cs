using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;

namespace GINtool
{

    public struct FC_BSU
    {
        public FC_BSU(double a, string b)
        {
            FC = a;
            BSU = b;
        }        
        public double FC { get; }
        public string BSU { get; }               
    }

    internal class Boundaries
    {
        struct low_high
        {
            public double low { get; set; }
            public double high { get; set; }
            public override string ToString()
            {
                return string.Format("{0},{1}", low, high);
            }
        }

        List<low_high> mLH = new List<low_high>();

        public Boundaries(string value)
        {
            
            mLH.Clear();
            if (value.Length > 0)
            {
                double[] arr = value.Split(',').Select(s => Double.Parse(s)).ToArray();
                for (int i = 0; i < arr.Length; i += 2)
                {
                    low_high lh = new low_high();
                    lh.low = arr[i];
                    lh.high = arr[i + 1];
                    mLH.Add(lh);
                }
            }
        }

        public string Store()
        {
            string value = String.Join(",", mLH.Select(i => i.ToString()).ToArray());
            return value;
        }
    }

    

    public partial class GinRibbon
    {

        bool gUpDown = false;
        bool gColorCells = false;
        bool gOverallCol = false;

        bool gDenseOutput = true;
        SysData.DataTable gRefWB = null;
        SysData.DataTable gRefMeans = null;
        string[] gColNames = null;
        Excel.Application gApplication = null;

        static List<string> gAvailItems = null;
        static List<string> gUpItems = null;
        static List<string> gDownItems = null;

        Boundaries gBoundaries = null;

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
            SysData.DataTable dtUniqRecords = new SysData.DataTable();
            dtUniqRecords = dt.DefaultView.ToTable(true, Columns);
            return dtUniqRecords;
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
       

        private bool LoadData()
        {
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

                CreateTableStatistics();

            }
            gApplication.EnableEvents = true;
            return gRefWB != null ? true : false;
        }

        private void CreateTableStatistics()
        {
            List<string> lString = new List<string>();
            lString.Add(Properties.Settings.Default.referenceRegulon);
            SysData.DataTable lRegs = GetDistinctRecords(gRefWB,lString.ToArray());

            gRefMeans = new SysData.DataTable();

            int totNrRows = gRefWB.Rows.Count;

            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn countColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
            gRefMeans.Columns.Add(regColumn);
            gRefMeans.Columns.Add(countColumn);
            gRefMeans.Columns.Add(avgColumn);

            foreach (SysData.DataRow lRow in lRegs.Rows)
            {
                string lVal = lRow[Properties.Settings.Default.referenceRegulon].ToString();
                int cnt = gRefWB.Select(string.Format("{0}='{1}'", Properties.Settings.Default.referenceRegulon, lVal)).Length;
                SysData.DataRow nRow = gRefMeans.Rows.Add();
                nRow["Regulon"] = lVal;
                nRow["Count"] = cnt;
                nRow["Average"] = ((double)cnt)/totNrRows;
            }                
        }

        private void GinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            gApplication = Globals.ThisAddIn.GetExcelApplication();
            lbRefFileName.Label = Properties.Settings.Default.referenceFile;

            gAvailItems = propertyItems("directionMapUnassigned");
            gUpItems = propertyItems("directionMapUp");
            gDownItems = propertyItems("directionMapDown");

            btApply.Enabled = false;
            ddBSU.Enabled = false;
            ddRegulon.Enabled = false;
            ddDir.Enabled = false;
            EnableOutputOptions(false);

            gBoundaries = new Boundaries(Properties.Settings.Default.fcBoundaries);

            btLoad.Enabled = System.IO.File.Exists(Properties.Settings.Default.referenceFile);


        }

        private Excel.Range GetActiveCell()
        {
            if (gApplication!=null)
            {
                return (Excel.Range) gApplication.Selection;
            }
            return null;
        }

        private void EnableOutputOptions(bool enable)
        {
            //cbOverall.Enabled = enable;
            //cbUpDown.Enabled = enable;
            //cbColor.Enabled = enable;
            tglDense.Enabled = enable;
        }


        private Excel.Worksheet GetActiveShet()
        {
            if (gApplication != null)
            {
                return (Excel.Worksheet)gApplication.ActiveSheet;
            }
            return null;
        }

        private void ConditionFormatRange(Excel.Range columnRange)
        {
            
            Excel.FormatConditions fcs = columnRange.FormatConditions;
            var formatCondition = fcs.Add(Excel.XlFormatConditionType.xlDatabar);

            formatCondition.MinPoint.Modify(Excel.XlConditionValueTypes.xlConditionValueAutomaticMin);
            formatCondition.MaxPoint.Modify(Excel.XlConditionValueTypes.xlConditionValueAutomaticMax);


            formatCondition.BarFillType = Excel.XlGradientFillType.xlGradientFillPath;
            formatCondition.Direction = Excel.Constants.xlContext;
            formatCondition.NegativeBarFormat.ColorType = Excel.XlDataBarNegativeColorType.xlDataBarColor;

            formatCondition.BarColor.Color = 8700771; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            formatCondition.BarColor.TintAndShade = 0; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);


            formatCondition.BarBorder.Color.Color = 8700771; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            formatCondition.BarBorder.Type = Excel.XlDataBarBorderType.xlDataBarBorderSolid;

            formatCondition.NegativeBarFormat.BorderColorType = Excel.XlDataBarNegativeColorType.xlDataBarColor;
            formatCondition.NegativeBarFormat.Parent.BarBorder.Type = Excel.XlDataBarBorderType.xlDataBarBorderSolid;

            formatCondition.AxisPosition = Excel.XlDataBarAxisPosition.xlDataBarAxisAutomatic;

            formatCondition.AxisColor.Color = 0; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            formatCondition.AxisColor.TintAndShade = 0;

            formatCondition.NegativeBarFormat.Color.Color = 255; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
            formatCondition.NegativeBarFormat.Color.TintAndShade = 0;

            formatCondition.NegativeBarFormat.BorderColor.Color = 255; // System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
            formatCondition.NegativeBarFormat.BorderColor.TintAndShade = 0;
        }

        private List<BsuRegulons> QueryResultTable(Excel.Range theCells)
        {            

            int nrRows = theCells.Rows.Count;                     
            List<BsuRegulons> lList = new List<BsuRegulons>();

            foreach (Excel.Range c in theCells.Rows)
            {
                object[,] value = null;
                BsuRegulons lMap = null;
                if (c.Columns.Count == 2)
                {
                    value = c.Value2;
                    lMap = new BsuRegulons((double)value[1, 1], value[1, 2].ToString());                    
                }

                else // only 1 column is selected                
                {
                    value = c.Value2;
                    lMap = new BsuRegulons(value[1, 1].ToString());
                }


                int rnr = 0;
                //int iR1 = c.Row;
                int rc = 0;
                int nUp = 0;
                int nDown = 0;

                if (value != null) 
                {
                    SysData.DataRow[] results = Lookup(lMap.BSU);

                    if (results.Length > 0)
                    {
                        for (int r = 0; r < results.Length; r++)
                        {
                            string item = results[r][Properties.Settings.Default.referenceRegulon].ToString();
                            string direction = results[r][Properties.Settings.Default.referenceDIR].ToString();

                            if (item.Length > 0) // loop over found regulons
                            {
                                lMap.REGULONS.Add(item);
                                rc += 1;
                                if (gUpItems.Contains(direction))
                                    nUp += 1;
                                if (gDownItems.Contains(direction))
                                    nDown += 1;
                            }
                        }
                    }

                    lMap.NRUP = nUp;
                    lMap.NRDOWN = nDown;
                    lMap.NET = nUp - nDown;
                    lMap.TOT = rc;               
                }
                else
                {                
                    lMap.NRUP = 0;
                    lMap.NRDOWN = 0;
                    lMap.NET = 0;
                    lMap.TOT = 0;
                }                

                lList.Add(lMap);
            }

            return lList;

        }
        private void ClearRange(Excel.Range range)
        {
            range.Clear();

            range.Interior.Pattern = Excel.Constants.xlNone;
            range.Interior.TintAndShade = 0;
            range.Interior.PatternTintAndShade = 0;
        }

        private void ClearOutputRange(Excel.Range theCells)
        {
            Excel.Worksheet theSheet = GetActiveShet();
            int nrRows = theCells.Rows.Count;
            int startC = theCells.Column;
            int startR = theCells.Row;
            int offsetColumn = startC + 2;
            int maxnrCols = 16384;
            int maxnrRows = 1048576;

            if ((nrRows + 1) > maxnrRows)
                nrRows -= 1;          

            Excel.Range tmpRange_ = null;
            if (gDenseOutput == false)
                tmpRange_ = (Excel.Range)theSheet.Range[theSheet.Cells[startR, offsetColumn], theSheet.Cells[startR + nrRows, maxnrCols - offsetColumn]];
            else
                tmpRange_ = (Excel.Range)theSheet.Range[theSheet.Cells[startR, offsetColumn], theSheet.Cells[startR + nrRows + 1, maxnrCols - offsetColumn]];
            tmpRange_.Clear();

            ClearRange(tmpRange_);
        }

        SysData.DataTable PrepareResultTable(List<BsuRegulons> lResults, bool bDense)
        {
            SysData.DataTable myTable = new System.Data.DataTable("mytable");
           
            if (bDense)
            {
                int maxcol = lResults[0].REGULONS.Count;

                // count max number of columns neccesary
                for (int r = 1; r < lResults.Count; r++)
                    if (maxcol < lResults[r].REGULONS.Count)
                        maxcol = lResults[r].REGULONS.Count;

                // add count column
                SysData.DataColumn countCol = new SysData.DataColumn("count_col",Type.GetType("System.Int16"));
                myTable.Columns.Add(countCol);

                // add variable columns
                for (int c = 0; c < maxcol; c++)
                {
                    SysData.DataColumn newCol = new SysData.DataColumn(string.Format("col_{0}", c+1));
                    myTable.Columns.Add(newCol);
                }

                // fill data
                for (int r=0;r<lResults.Count;r++)
                {
                    SysData.DataRow newRow = myTable.Rows.Add();
                    newRow["count_col"]= lResults[r].TOT;
                    for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                        newRow[string.Format("col_{0}",c+1)] = lResults[r].REGULONS[c];
                }

            }
            else // generate sparse output
            {
                List<string> allRegs = new List<string>();
                for (int r = 0; r < lResults.Count; r++)
                    allRegs.AddRange(lResults[r].REGULONS);

                List<string> uRegs = allRegs.Distinct().ToList();
                
                // add count column
                SysData.DataColumn countCol = new SysData.DataColumn("count_col", Type.GetType("System.Int16"));
                myTable.Columns.Add(countCol);

                // add variable columns
                for (int c = 0; c < uRegs.Count; c++)
                {
                    SysData.DataColumn newCol = new SysData.DataColumn(string.Format(uRegs[c]));
                    myTable.Columns.Add(newCol);
                }

                // fill data
                for (int r = 0; r < lResults.Count; r++)
                {
                    SysData.DataRow newRow = myTable.Rows.Add();
                    newRow["count_col"] = lResults[r].TOT;
                    for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                        newRow[lResults[r].REGULONS[c]] = lResults[r].REGULONS[c];
                }

                // now reorder the output
                int[] nrTypes = new int[myTable.Columns.Count-1];

                for (int c = 1; c < myTable.Columns.Count; c++)
                {
                    int cc = 0;
                    for (int r = 0; r < myTable.Rows.Count; r++)
                    {
                        string item = myTable.Rows[r].ItemArray[c].ToString();
                        if (item.Length > 0)
                            cc += 1;
                    }
                    nrTypes[c-1] = cc;
                }

                string[] keys = uRegs.ToArray();

                // Sort with keys only works in ascending order
                Array.Sort(nrTypes, keys);
                Array.Reverse(nrTypes);
                Array.Reverse(keys);
                
                for (int c = 0; c < myTable.Columns.Count-1; c++)
                    myTable.Columns[keys[c]].SetOrdinal(c+1);
                
                SysData.DataRow _newRow = myTable.Rows.Add();
                for (int c = 0; c < myTable.Columns.Count - 1; c++)
                    _newRow[c + 1] = nrTypes[c];

            }
            return myTable;
        }

        private List<FC_BSU> GenerateOutput(bool bDense)
        {
            Excel.Range theInputCells = GetActiveCell();
            Excel.Worksheet theSheet = GetActiveShet();
           
            int nrRows = theInputCells.Rows.Count;            
            int startC = theInputCells.Column;
            int startR = theInputCells.Row;
            int offsetColumn = startC + 2;            
                      
            if (theInputCells.Columns.Count != 2)
            {
                MessageBox.Show("Please select 2 columns, first FC, second BSU");
                return null;
            }

            ClearOutputRange(theInputCells);          

            // generate the results for outputting the data and summary
            List<BsuRegulons> lResults=QueryResultTable(theInputCells);
            // output the data
            SysData.DataTable lOut = PrepareResultTable(lResults, bDense);
            FastDtToExcel(lOut,theSheet,startR,offsetColumn,startR+nrRows-(bDense?1:0), offsetColumn+lOut.Columns.Count-1);


            List<FC_BSU> lTable = new List<FC_BSU>();

            for (int r = 0; r < nrRows; r++)
                for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                    lTable.Add(new FC_BSU(lResults[r].FC, lResults[r].REGULONS[c]));             
                                              
            return lTable;
        }

        private void FastDtToExcel(System.Data.DataTable dt, Excel.Worksheet sheet, int firstRow, int firstCol, int lastRow, int lastCol)
        {
            Excel.Range top = sheet.Cells[firstRow, firstCol];
            Excel.Range bottom = sheet.Cells[lastRow, lastCol];
            Excel.Range all = (Excel.Range)sheet.get_Range(top, bottom);
            object[,] arrayDT = new object[dt.Rows.Count, dt.Columns.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
                for (int j = 0; j < dt.Columns.Count; j++)
                    arrayDT[i, j] = dt.Rows[i][j]; //.ToString();
            all.Value2 = arrayDT;
        }

       
        private void btApply_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;
            
            List<FC_BSU> lOutput= GenerateOutput(gDenseOutput);

            //CreateStatisticsSheet(lOutput);

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }

        private void CreateStatisticsSheet(List<string> aAllItems)
        {
            SysData.DataTable table = new SysData.DataTable();
            //using (var reader = FM.ObjectReader.Create(aAllItems))
            //{
            //    table.Load(reader);
            //}



            //throw new NotImplementedException();
        }


        private SysData.DataTable CreateUsageTable(SysData.DataTable aTable)
        {
            
            SysData.DataTable lTable  = new SysData.DataTable();
            List<string> lString = new List<string>();
            lString.Add("Regulon");
            SysData.DataTable lRegs = GetDistinctRecords(aTable, lString.ToArray());
            
            int totNrRows = aTable.Rows.Count;

            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn countColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
            SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
            SysData.DataColumn low1Column = new SysData.DataColumn("low1", Type.GetType("System.Double"));
            SysData.DataColumn low2Column = new SysData.DataColumn("low2", Type.GetType("System.Double"));
            SysData.DataColumn low3Column = new SysData.DataColumn("low3", Type.GetType("System.Double"));
            SysData.DataColumn low4Column = new SysData.DataColumn("low4", Type.GetType("System.Double"));
            SysData.DataColumn high1Column = new SysData.DataColumn("high1", Type.GetType("System.Double"));
            SysData.DataColumn high2Column = new SysData.DataColumn("high2", Type.GetType("System.Double"));
            SysData.DataColumn high3Column = new SysData.DataColumn("high3", Type.GetType("System.Double"));
            SysData.DataColumn high4Column = new SysData.DataColumn("high4", Type.GetType("System.Double"));
            
            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(countColumn);
            lTable.Columns.Add(avgColumn);
            lTable.Columns.Add(low1Column);
            lTable.Columns.Add(low2Column);
            lTable.Columns.Add(low3Column);
            lTable.Columns.Add(low4Column);
            lTable.Columns.Add(high1Column);
            lTable.Columns.Add(high2Column);
            lTable.Columns.Add(high3Column);
            lTable.Columns.Add(high4Column);


            foreach (SysData.DataRow lRow in lRegs.Rows)
            {
                string lVal = lRow[Properties.Settings.Default.referenceRegulon].ToString();
                int cnt = gRefWB.Select(string.Format("{0}='{1}'", Properties.Settings.Default.referenceRegulon, lVal)).Length;
                SysData.DataRow nRow = lTable.Rows.Add();
                nRow["Regulon"] = lVal;
                nRow["Count"] = cnt;
                nRow["Average"] = ((double)cnt) / totNrRows;
                //nRow["grp1"]





            }
            return lTable;
        }



        private void tglDense_Click(object sender, RibbonControlEventArgs e)
        {
            gDenseOutput = tglDense.Checked == false;
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
            int nrS = excelworkBook.Sheets.Count;
            // Set workbook to first worksheet
            Excel.Worksheet ws = (Excel.Worksheet)excelworkBook.Sheets[1];
            Properties.Settings.Default.referenceSheetName = ws.Name;

          
            excelworkBook.Close();

            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
        }


        private void Fill_DropDownBoxes()
        {
            gApplication.EnableEvents = false;

            ddBSU.Items.Clear();
            ddRegulon.Items.Clear();

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

            ddBSU.Enabled = true;
            ddRegulon.Enabled = true;
            ddDir.Enabled = true;
            btRegDirMap.Enabled = true;
            gApplication.EnableEvents = true;

        }

        private void btLoad_Click(object sender, RibbonControlEventArgs e)
        {
            
            gApplication.EnableEvents = false;
            if (LoadData())
            {                
                Fill_DropDownBoxes();
                if (gDownItems.Count == 0 && gUpItems.Count == 0 && gAvailItems.Count == 0)
                    LoadDirectionOptions();
                btApply.Enabled = true;
                EnableOutputOptions(true);                
            }
            gApplication.EnableEvents = true;
        }

        private void EnableItems(bool enable)
        {
            btLoad.Enabled = enable;
            ddBSU.Enabled = enable;
            ddRegulon.Enabled = enable;

        }

        private void btSelectFile_Click(object sender, RibbonControlEventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {

                    Properties.Settings.Default.referenceFile = openFileDialog.FileName;
                    lbRefFileName.Label = Properties.Settings.Default.referenceFile;
                    load_Worksheets();
                    btLoad.Enabled = true;                    

                }
            }
        }

        private void ddBSU_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceBSU = ddBSU.SelectedItem.Label;
        }

        private void ddRegulon_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceRegulon = ddRegulon.SelectedItem.Label;
        }

        private void btRegDirMap_Click(object sender, RibbonControlEventArgs e)
        {           
            dlgUpDown dlgUD = new dlgUpDown(gAvailItems, gUpItems, gDownItems);            
            dlgUD.ShowDialog();

            storeValue("directionMapUnassigned", gAvailItems);
            storeValue("directionMapUp", gUpItems);
            storeValue("directionMapDown", gDownItems);

        }

        private void ddDir_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.referenceDIR = ddDir.SelectedItem.Label;
            gAvailItems.Clear();
            gUpItems.Clear();
            gDownItems.Clear();
            LoadDirectionOptions();
        }

        private void cbDense_Click(object sender, RibbonControlEventArgs e)
        {
            //gDenseOutput = cbDense.Checked == false;
        }
       
    }
}
