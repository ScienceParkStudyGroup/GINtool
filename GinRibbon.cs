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

    public partial class GinRibbon
    {

        bool gGenReport = false;
        bool gDenseOutput = true;
        SysData.DataTable gRefWB = null;
        SysData.DataTable gRefStats = null;
        string[] gColNames = null;
        Excel.Application gApplication = null;

        static List<string> gAvailItems = null;
        static List<string> gUpItems = null;
        static List<string> gDownItems = null;


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

                // generate database frequency table
                CreateTableStatistics();

            }
            gApplication.EnableEvents = true;
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
            lbRefFileName.Label = Properties.Settings.Default.referenceFile;

            gAvailItems = propertyItems("directionMapUnassigned");
            gUpItems = propertyItems("directionMapUp");
            gDownItems = propertyItems("directionMapDown");

            btApply.Enabled = false;
            ddBSU.Enabled = false;
            ddRegulon.Enabled = false;
            ddDir.Enabled = false;
            EnableOutputOptions(false);

            gDenseOutput = true;
            cbReport.Checked = false;            

            btLoad.Enabled = System.IO.File.Exists(Properties.Settings.Default.referenceFile);


        }

        private Excel.Range GetActiveCell()
        {
            if (gApplication != null)
            {
                return (Excel.Range)gApplication.Selection;
            }
            return null;
        }

        private void EnableOutputOptions(bool enable)
        {
            ebLow.Enabled = enable;
            ebMid.Enabled = enable;
            ebHigh.Enabled = enable;
            cbReport.Enabled = enable;
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

            double lowVal = Properties.Settings.Default.fcLOW;

            //int nrRows = theCells.Rows.Count;                     
            List<BsuRegulons> lList = new List<BsuRegulons>();

            foreach (Excel.Range c in theCells.Rows)
            {
                string lBSU;
                bool hasFC = false;
                double lFC = 0;
                BsuRegulons lMap;
                if (c.Columns.Count == 2)
                {
                    object[,] value = c.Value2;
                    lFC = (double)value[1, 1];
                    lBSU = value[1, 2].ToString();
                    lMap = new BsuRegulons(lFC, lBSU);
                    hasFC = true;

                }

                else // only 1 column is selected                
                {
                    object value = c.Value2;
                    lMap = new BsuRegulons(value.ToString());
                    lBSU = value.ToString();
                }

                if (lBSU.Length > 0)
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

            return lList;

        }
        private void ClearRange(Excel.Range range)
        {            
            range.Interior.Pattern = Excel.Constants.xlNone;
            range.Interior.TintAndShade = 0;
            range.Interior.PatternTintAndShade = 0;
            range.Clear();
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

            Excel.Range tmpRange_;
            if (gDenseOutput == false)
                tmpRange_ = (Excel.Range)theSheet.Range[theSheet.Cells[startR, offsetColumn], theSheet.Cells[startR + nrRows, maxnrCols - offsetColumn]];
            else
                tmpRange_ = (Excel.Range)theSheet.Range[theSheet.Cells[startR, offsetColumn], theSheet.Cells[startR + nrRows + 1, maxnrCols - offsetColumn]];
            tmpRange_.Clear();

            ClearRange(tmpRange_);
        }

        (SysData.DataTable, SysData.DataTable) PrepareResultTable(List<BsuRegulons> lResults, bool bDense)
        {
            SysData.DataTable myTable = new System.Data.DataTable("mytable");
            SysData.DataTable clrTable = new System.Data.DataTable("colortable");

            if (bDense)
            {
                int maxcol = lResults[0].REGULONS.Count;

                // count max number of columns neccesary
                for (int r = 1; r < lResults.Count; r++)
                    if (maxcol < lResults[r].REGULONS.Count)
                        maxcol = lResults[r].REGULONS.Count;

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
                int[] nrTypes = new int[myTable.Columns.Count - 1];

                for (int c = 1; c < myTable.Columns.Count; c++)
                {
                    int cc = 0;
                    for (int r = 0; r < myTable.Rows.Count; r++)
                    {
                        string item = myTable.Rows[r].ItemArray[c].ToString();
                        if (item.Length > 0)
                            cc += 1;
                    }
                    nrTypes[c - 1] = cc;
                }

                string[] keys = uRegs.ToArray();

                // Sort with keys only works in ascending order
                Array.Sort(nrTypes, keys);
                Array.Reverse(nrTypes);
                Array.Reverse(keys);

                for (int c = 0; c < myTable.Columns.Count - 1; c++)
                    myTable.Columns[keys[c]].SetOrdinal(c + 1);

                SysData.DataRow _newRow = myTable.Rows.Add();
                for (int c = 0; c < myTable.Columns.Count - 1; c++)
                    _newRow[c + 1] = nrTypes[c];

                return (myTable, null);
            }
        }

        private (List<FC_BSU>, List<BsuRegulons>) GenerateOutput(bool bDense)
        {
            Excel.Range theInputCells = GetActiveCell();
            Excel.Worksheet theSheet = GetActiveShet();

            int nrRows = theInputCells.Rows.Count;
            int startC = theInputCells.Column;
            int startR = theInputCells.Row;

            bool genSummary = theInputCells.Columns.Count == 2;
            int offsetColumn = startC + (genSummary ? 2 : 1);


            if (gGenReport)
            {
                if (theInputCells.Columns.Count != 2)
                {
                    MessageBox.Show("Please select 2 columns (first FC, second BSU)");
                    return (null,null);
                }
            }
            else
            {
                if (theInputCells.Columns.Count != 1)
                {
                    MessageBox.Show("Please select 1 column with BSU entries");
                    return (null,null);
                }
            }


            ClearOutputRange(theInputCells);

            // generate the results for outputting the data and summary
            List<BsuRegulons> lResults = QueryResultTable(theInputCells);
            // output the data
            var lOut = PrepareResultTable(lResults, bDense);
            SysData.DataTable lTable = lOut.Item1;
            SysData.DataTable clrTbl;


            FastDtToExcel(lTable, theSheet, startR, offsetColumn, startR + nrRows - (bDense ? 1 : 0), offsetColumn + lTable.Columns.Count - 1);
            
            if (bDense)
            {
                clrTbl = lOut.Item2;
                ColorCells(clrTbl, theSheet, startR, offsetColumn + 1, startR + nrRows - (bDense ? 1 : 0), offsetColumn + lTable.Columns.Count - 1);
            }


            if (genSummary)
            {
                List<FC_BSU> lOutput = new List<FC_BSU>();

                for (int r = 0; r < nrRows; r++)
                    for (int c = 0; c < lResults[r].REGULONS.Count; c++)
                    {
                        int val = 0;
                        if (lResults[r].UP.Contains(c))
                            val = 1;
                        if (lResults[r].DOWN.Contains(c))
                            val = -1;

                        lOutput.Add(new FC_BSU(lResults[r].FC, lResults[r].REGULONS[c], val));
                    }


                return (lOutput, lResults);
            }
            else
                return (null,null);
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

            int col = 1;

            Excel.Range top = lNewSheet.Cells[1, 4];
            Excel.Range bottom = lNewSheet.Cells[1, 11];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Counts";
            all.HorizontalAlignment = Excel.Constants.xlCenter;

            top = lNewSheet.Cells[1, 13];
            bottom = lNewSheet.Cells[1, 20];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "Percentage";
            all.HorizontalAlignment = Excel.Constants.xlCenter;

            top = lNewSheet.Cells[1, 21];
            bottom = lNewSheet.Cells[1, 22];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.Merge();
            all.Value = "REGULATIONS";
            all.HorizontalAlignment = Excel.Constants.xlCenter;

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

            top = lNewSheet.Cells[3, 13];
            bottom = lNewSheet.Cells[2 + theTable.Rows.Count, 22];
            all = (Excel.Range)lNewSheet.get_Range(top, bottom);
            all.NumberFormat = "###%";

        }

        private SysData.DataTable ReformatResults(List<FC_BSU> aList)
        {
            // find unique regulons

            SysData.DataTable lTable = new SysData.DataTable("FC_BSU");
            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Single"));
            SysData.DataColumn dirColumn = new SysData.DataColumn("DIR", Type.GetType("System.Int32"));


            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(fcColumn);
            lTable.Columns.Add(dirColumn);

            for (int r = 0; r < aList.Count; r++)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                lRow["Regulon"] = aList[r].BSU;
                lRow["FC"] = aList[r].FC;
                lRow["DIR"] = aList[r].DIR;
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


        private void CreateCombinedSheet(SysData.DataTable aTable)
        {
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();

           
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

            lNewSheet.Cells[1, col++] = "FC";
            lNewSheet.Cells[1, col++] = "BSU";
            
          
            int maxRegulons = 0;

            for (int r = 0; r < aTable.Rows.Count; r++)
            {
                SysData.DataRow clrRow = aTable.Rows[r];
                for (int c = 0; c < clrRow.ItemArray.Length; c++)
                {
                    Excel.Range lR = all.Cells[r + 2, c + 1];
                    int UpPos = clrRow[c].ToString().IndexOf("#");
                    int DownPos = clrRow[c].ToString().IndexOf("@");                                        

                    lR.Value = clrRow[c];

                    if (clrRow[c].ToString().Length == 0)
                        continue;

                    if (maxRegulons < c)
                        maxRegulons = c;

                    if (UpPos == -1 && DownPos == -1)
                        continue;

                    if (UpPos > 0)
                    {
                        Excel.Characters lChar = lR.Characters[UpPos+1, 1];
                        lChar.Text = "á"; // the arrow up
                        lChar.Font.Name = "Wingdings";
                        lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {

                        Excel.Characters lChar = lR.Characters[DownPos + 1, 1];
                        lChar.Text = "â"; // the arrow down
                        lChar.Font.Name = "Wingdings";
                        lR.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
                    }                                            
                }
            }


            for (int c = 0; c < (maxRegulons - 1); c++)
                lNewSheet.Cells[1, col++] = string.Format("Regulon_{0}", c + 1);


            all.Columns.AutoFit();

            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;

        }


        private SysData.DataTable CreateCombinedTable(SysData.DataTable aUsageTbl, List <BsuRegulons> lLst)
        {
            SysData.DataTable lTable = new SysData.DataTable();

            //col = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            //lTable.Columns.Add(col);
            SysData.DataColumn  col = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            lTable.Columns.Add(col);
            col = new SysData.DataColumn("BSU", Type.GetType("System.String"));
            lTable.Columns.Add(col);


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

                if (Math.Abs(lLst[r].FC) > lowVal)
                {
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["FC"] = lLst[r].FC;
                    lRow["BSU"] = lLst[r].BSU;
                    for (int i = 0; i < lLst[r].REGULONS.Count; i++)
                    {

                        double nrUP = 0, nrDOWN = 0, percUP = 0, percDOWN = 0;
                        SysData.DataRow[] lHit = aUsageTbl.Select(string.Format("Regulon = '{0}'", lLst[r].REGULONS[i]));
                        nrUP = Double.Parse(lHit[0]["nr_UP"].ToString());
                        nrDOWN = Double.Parse(lHit[0]["nr_DOWN"].ToString());

                        Double.TryParse(lHit[0]["perc_UP"].ToString(),out percUP);
                        Double.TryParse(lHit[0]["perc_DOWN"].ToString(),out percDOWN);

                        string lVal = "";
                        if (nrUP>nrDOWN || nrUP==nrDOWN)
                            lVal = percUP.ToString("P0") + "# " + nrUP.ToString("P0") + "-tot";
                        else
                            lVal = percDOWN.ToString("P0") + "@ " + nrDOWN.ToString("P0") + "-tot";                                                 

                        lRow[string.Format("Regulon_{0}", i + 1)] = lLst[r].REGULONS[i] + " " + lVal;

                    }
                }
            }


            return lTable;
        }

        private (SysData.DataTable, SysData.DataTable) CreateUsageTable(List<FC_BSU> aList)
        {
            SysData.DataTable _fc_BSU = ReformatResults(aList);

            SysData.DataTable lTable = new SysData.DataTable();
            SysData.DataTable lTableCombine = new SysData.DataTable();


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


            col = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            lTableCombine.Columns.Add(col);
            
            col = new SysData.DataColumn("perc_DOWN", Type.GetType("System.Double"));
            lTableCombine.Columns.Add(col);
            col = new SysData.DataColumn("perc_UP", Type.GetType("System.Double"));
            lTableCombine.Columns.Add(col);


            col = new SysData.DataColumn("nr_DOWN", Type.GetType("System.Double"));
            lTableCombine.Columns.Add(col);
            col = new SysData.DataColumn("nr_UP", Type.GetType("System.Double"));
            lTableCombine.Columns.Add(col);


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

                lNewRow["totrel"] = nrTOT;

                if (nrTOT > 0)
                {
                    lNewRow["perc_DOWN"] = (double)nrDOWN / (double)(nrTOT);
                    lNewRow["perc_UP"] = (double)nrUP / (double)(nrTOT);
                }


                double lCount = (double)_tmp.Length;

                lNewRow = lTableCombine.Rows.Add();
                lNewRow["Regulon"] = reg;

                
                if (nrTOT > 0)
                {
                    lNewRow["perc_DOWN"] = (double)nrDOWN / (double)(nrTOT);
                    lNewRow["perc_UP"] = (double)nrUP / (double)(nrTOT);
                }

                lNewRow["nr_DOWN"] = ((double)nrDOWN)/lCount;
                lNewRow["nr_UP"] = ((double)nrUP)/lCount;


            }

            SysData.DataView dv = lTable.DefaultView;
            dv.Sort = "totrel desc";

            return (dv.ToTable(),lTableCombine);
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
                LoadFCDefaults();
                EnableOutputOptions(true);
            }
            gApplication.EnableEvents = true;
        }

        private void LoadFCDefaults()
        {
            ebLow.Text = Properties.Settings.Default.fcLOW.ToString();
            ebMid.Text = Properties.Settings.Default.fcMID.ToString();
            ebHigh.Text = Properties.Settings.Default.fcHIGH.ToString();
        }

        private void EnableItems(bool enable)
        {
            btLoad.Enabled = enable;
            ddBSU.Enabled = enable;
            ddRegulon.Enabled = enable;

        }

        private void btSelectFile_Click(object sender, RibbonControlEventArgs e)
        {
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
    

        private void cbReport_Click(object sender, RibbonControlEventArgs e)
        {
            gGenReport = cbReport.Checked;
        }

        private void btApply_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            (List<FC_BSU> lOutput, List<BsuRegulons> lList) = GenerateOutput(gDenseOutput);

            if (lOutput != null & gGenReport)
            {
                (SysData.DataTable lSummary, SysData.DataTable lCombineInfo) = CreateUsageTable(lOutput);
                CreateSummarySheet(lSummary);
                SysData.DataTable lCombined = CreateCombinedTable(lCombineInfo, lList);
                CreateCombinedSheet(lCombined);
            }

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }
    }


    public struct FC_BSU
    {
        public FC_BSU(double a, string b, int dir)
        {
            FC = a;
            BSU = b;
            DIR = dir;
        }
        public double FC { get; }
        public string BSU { get; }
        public double DIR { get; }
    }


}
