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
        SysData.DataTable gRefWB = null;
        SysData.DataTable gRefStats = null;
        SysData.DataTable gRefOperons = null;
        string[] gColNames = null;
        Excel.Application gApplication = null;

        bool gCompositPlot = false;
        bool gQPlot = false;

        static List<string> gAvailItems = null;
        static List<string> gUpItems = null;
        static List<string> gDownItems = null;

        List<int> gExcelErrorValues = null;

        EnrichmentAnalysis enrichmentAnalysis1 = null;

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




        private bool LoadOperonData()
        {
            if (Properties.Settings.Default.operonFile.Length == 0 || Properties.Settings.Default.operonSheet.Length == 0)
                return false;

            gApplication.EnableEvents = false;
            
            SysData.DataTable _tmp = ExcelUtils.ReadExcelToDatable(gApplication, Properties.Settings.Default.operonSheet, Properties.Settings.Default.operonFile, 1, 1);
            gRefOperons = new SysData.DataTable("OPERONS");
            gRefOperons.CaseSensitive = false;
            gRefOperons.Columns.Add("operon", Type.GetType("System.String"));
            gRefOperons.Columns.Add("gene", Type.GetType("System.String"));

            foreach(SysData.DataRow lRow in _tmp.Rows)
            {
                string[] lItems = lRow.ItemArray[0].ToString().Split('-');
                for (int i = 0; i < lItems.Length; i++)
                {
                    SysData.DataRow lNewRow = gRefOperons.Rows.Add();
                    lNewRow["operon"] = lItems[0];
                    lNewRow["gene"] = lItems[i];
                }
            }
            gApplication.EnableEvents = true;
            return gRefOperons.Rows.Count>0;
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
            btnRegulonFileName.Label = Properties.Settings.Default.referenceFile;
            btnOperonFile.Label = Properties.Settings.Default.operonFile;

            if (Properties.Settings.Default.operonFile.Length == 0)
                btnOperonFile.Label = "No file selected";

            gAvailItems = propertyItems("directionMapUnassigned");
            gUpItems = propertyItems("directionMapUp");
            gDownItems = propertyItems("directionMapDown");

            btApply.Enabled = false;
            ddBSU.Enabled = false;
            ddGene.Enabled = false;
            ddRegulon.Enabled = false;
            ddDir.Enabled = false;

            edtMaxGroups.Enabled = false;
            btnPalette.Enabled = false;

            cbComposit.Checked = Properties.Settings.Default.compositPlot;
            cbQplot.Checked = Properties.Settings.Default.qPlot;
            gCompositPlot = cbComposit.Checked;
            gQPlot = cbQplot.Enabled;

            if (cbComposit.Checked || cbQplot.Enabled)
                if (enrichmentAnalysis1 == null)
                {
                    enrichmentAnalysis1 = new EnrichmentAnalysis(gApplication);
                }

            EnableOutputOptions(false);

            gExcelErrorValues = ((int[])Enum.GetValues(typeof(ExcelUtils.CVErrEnum))).ToList();

            if (Properties.Settings.Default.use_pvalues)
            {
                splitButton3.Label = but_pvalues.Label;
                splitButton3.Image = but_pvalues.Image;
            }
            else
            {
                splitButton3.Label = but_fc.Label;
                splitButton3.Image = but_pvalues.Image;
            }

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
            splitButton3.Enabled = enable;
            splitbtnEA.Enabled = enable;
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

        private (List<FC_BSU>, List<BsuRegulons>) GenerateOutput()
        {
            Excel.Range theInputCells = GetActiveCell();
            Excel.Worksheet theSheet = GetActiveShet();

            if (theSheet == null)
                return (null,null);

            if (theSheet.Name.Contains("Plot_"))
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
                

            int nrRows = theInputCells.Rows.Count;
            int startC = theInputCells.Column;
            int startR = theInputCells.Row;

            // from now always assume 3 columns.. p-value, fc, bsu
            int offsetColumn = 1;

            if(theInputCells.Columns.Count !=3)
            {
                MessageBox.Show("Please select 3 columns (first P-Value, second FC, third BSU)");
                return (null, null);

            }                       
            // generate the results for outputting the data and summary
            List<BsuRegulons> lResults = QueryResultTable(theInputCells);
            // output the data
            var lOut = PrepareResultTable(lResults);
            SysData.DataTable lTable = lOut.Item1;
            SysData.DataTable clrTbl;


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
            Excel.Range bottom = lNewSheet.Cells[lTable.Rows.Count+1, lTable.Columns.Count];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();
          
            clrTbl = lOut.Item2;
            ColorCells(clrTbl, lNewSheet, startR, offsetColumn + 5, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);
                     
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


            return (lOutput, lResults);
           
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
                if(sheet is Excel.Chart)
                    _sheets.Add(((Excel.Chart)sheet).Name);
                else
                    _sheets.Add(((Excel.Worksheet)sheet).Name);
            }

            return _sheets;
                
        }


        private void renameWorksheet(object aSheet, string wsBase)
        {
            // create a sheetname starting with wsBase
            List<string> currentSheets = listSheets();
            int s = 1;
            while (currentSheets.Contains(string.Format("{0}_{1}", wsBase, s)))
                s += 1;
            
            if(aSheet is Excel.Chart)
                ((Excel.Chart) aSheet).Name = string.Format("{0}_{1}", wsBase, s);
            else
                ((Excel.Worksheet)(aSheet)).Name = string.Format("{0}_{1}", wsBase, s);
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
            if(gOperonOutput)
                lNewSheet.Cells[1, col++] = "OPERON(S)";


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
            
            if (gOperonOutput)
            {
                col = new SysData.DataColumn("OPERON", Type.GetType("System.String"));
                lTable.Columns.Add(col);
            }

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

                    if (gRefOperons != null)
                    {
                        string lOperon = "";
                        if (lLst[r].GENE != "")
                        {
                            SysData.DataRow[] lOperons = gRefOperons.Select(string.Format("gene='{0}'", lLst[r].GENE));
                            List<string> strOperons = new List<string>();
                            for (int i = 0; i < lOperons.Length; i++)
                            {
                                strOperons.Add(lOperons[i]["operon"].ToString());
                            }

                            lOperon = String.Join(", ", strOperons.ToArray());
                        }

                        lRow["OPERON"] = lOperon;
                    }

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
                if(int.TryParse(_tmp2[0]["Count"].ToString(),out int totcount))
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
            edtMaxGroups.Enabled = true;
            btnPalette.Enabled = true;

            gApplication.EnableEvents = true;


        }

        private void btLoad_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;
            if (LoadData())
            {
                gOperonOutput = LoadOperonData();
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
            editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();
        }

        private void EnableItems(bool enable)
        {
            btLoad.Enabled = enable;
            ddBSU.Enabled = enable;
            ddRegulon.Enabled = enable;
            ddGene.Enabled = enable;
            edtMaxGroups.Enabled = enable;
            btnPalette.Enabled = enable;

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

        #region main_routine
        private void btApply_Click(object sender, RibbonControlEventArgs e)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            (List<FC_BSU> lOutput, List<BsuRegulons> lList) = GenerateOutput();

            if (lOutput != null)
            {
                (SysData.DataTable lSummary, SysData.DataTable lCombineInfo) = CreateUsageTable(lOutput);
                CreateSummarySheet(lSummary);
                SysData.DataTable lCombined = CreateCombinedTable(lCombineInfo, lList);
                CreateCombinedSheet(lCombined);
                if (gCompositPlot)
                    CreateCompositPlot(lOutput, lSummary);
                if (gQPlot)
                    CreateQPlot(lOutput, lSummary);
            }

            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
        }

        private void CreateCompositPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            //Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            //renameWorksheet(lNewSheet, "CompositPlot");

            SysData.DataTable _fc_BSU = ReformatResults(aOutput);


            HashSet<string> lRegulons = new HashSet<string>();

            //string.Format("[{0}] LIKE '%{1}%'", Properties.Settings.Default.referenceBSU, value

            SysData.DataView lRelevant = aSummary.AsDataView();
            lRelevant.RowFilter = "totrel>0";
            SysData.DataTable dataTable = lRelevant.ToTable(); 
               

            foreach(SysData.DataRow row in dataTable.Rows)
            {
                lRegulons.Add(row.ItemArray[0].ToString());
            }



            //for (int r = 0; r < lRelevant.Length; r++)
            //    lUnique.Add(lRelevant[r].["Regulon"]);
            string subsets = string.Join(",", lRegulons.ToArray());
            subsets = string.Join(",", subsets.Split(',').Select(x => $"'{x}'"));

            SysData.DataView dataView = _fc_BSU.AsDataView();
            dataView.RowFilter = String.Format("Regulon in ({0})", subsets);
            dataTable = dataView.ToTable();


            //FastDtToExcel(dataTable, lNewSheet, 1, 1, dataTable.Rows.Count,dataTable.Columns.Count);

            //SysData.DataRow[] geneTable = _fc_BSU.Select(String.Format("Regulon in ({0})", subsets));
            //Excel.Shape distPlot = enrichmentAnalysis1.DrawDistributionPlot(lRegulons, dataTable,);

            //distPlot.Name = "distributionPlot";
            //distPlot.Copy();

            //Excel.Range aRange = lNewSheet.Cells[1, 1];
            //lNewSheet.Paste(aRange);

            //foreach (Excel.Shape aShape in lNewSheet.Shapes)
            //{
            //    if (aShape.Name == "distributionPlot")
            //    {
            //        aShape.Top = 10;
            //        aShape.Left = 50;
            //        aShape.Width = 700;
            //        aShape.Height = 300;

            //    }
            //}

            

   

            //Excel.Worksheet lNewSheet = (Excel.Worksheet) 
             Excel.Chart aChart = enrichmentAnalysis1.CreateExcelChart(lRegulons, dataTable);
             renameWorksheet(aChart, "CompositPlot");

            //eaPlot.Name = "eaPlot";
            //eaPlot.Copy();

            //aRange = lNewSheet.Cells[1, 1];
            //lNewSheet.Paste(aRange);

            //foreach (Excel.Shape aShape in lNewSheet.Shapes)
            //{
            //    if (aShape.Name == "eaPlot")
            //    {
            //        aShape.Top = 310;
            //        aShape.Left = 50;
            //        aShape.Width = 400;
            //        //aShape.Height = 500;

            //    }
            //}





            gApplication.EnableEvents = true;
            gApplication.DisplayAlerts = true;
            //throw new NotImplementedException();
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

        }
       
        private void editMinPval_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (float.TryParse(editMinPval.Text, out float val))
            {
                // set the text value to what is parsed
                editMinPval.Text = val.ToString();                
                Properties.Settings.Default.pvalue_cutoff = val;                
            }
            else            
                editMinPval.Text = Properties.Settings.Default.pvalue_cutoff.ToString();                            
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            splitButton3.Label = but_pvalues.Label;
            splitButton3.Image = but_pvalues.Image;
            Properties.Settings.Default.use_pvalues = true;
        }

        private void but_fc_Click(object sender, RibbonControlEventArgs e)
        {
            splitButton3.Label = but_fc.Label;
            splitButton3.Image = but_fc.Image;
            Properties.Settings.Default.use_pvalues = false;
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
        }

        private void btnEA_Click(object sender, RibbonControlEventArgs e)
        {
            

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            renameWorksheet(lNewSheet, "Plots_");
            EnrichmentAnalysis enrichmentAnalysis = new EnrichmentAnalysis(gApplication);
            
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

        private void clrExcel_Click(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrExcel.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Excel;
        }

        private void clrGray_Click(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrGray.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Grayscale;

        }


        private void cbComposit_Click(object sender, RibbonControlEventArgs e)
        {
            gCompositPlot = cbComposit.Checked;
            Properties.Settings.Default.compositPlot = gCompositPlot;
            if (gCompositPlot)
                if (enrichmentAnalysis1 == null)
                {
                    enrichmentAnalysis1 = new EnrichmentAnalysis(gApplication);
                }
        }

        private void cbQplot_Click(object sender, RibbonControlEventArgs e)
        {
            gQPlot = cbQplot.Checked;
            Properties.Settings.Default.qPlot = gQPlot;
            
        }

        private void clrGray_Click_1(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrGray.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Grayscale;
        }

        private void clrBerry_Click(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrBerry.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Berry;
        }

        private void clrBright_Click(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrBright.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Bright;
        }

        private void clrBrightPastel_Click(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrBrightPastel.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.BrightPastel;
        }

        private void clrChocolate_Click(object sender, RibbonControlEventArgs e)
        {
            btnPalette.Image = clrChocolate.Image;
            Properties.Settings.Default.defaultPalette = (int)System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Chocolate;
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
