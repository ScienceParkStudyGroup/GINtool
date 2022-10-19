using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;

namespace GINtool
{
    public partial class GinRibbon
    { /// <summary>
      /// This routine reformats the original data to a data table that can be exported to an excel sheet. The second data table contains the color formatting of the different cells.
      /// </summary>
      /// <param name="lResults"></param>
      /// <returns></returns>
        (SysData.DataTable, SysData.DataTable) PrepareResultTable(List<BsuLinkedItems> lResults)
        {
            SysData.DataTable myTable = new System.Data.DataTable("mytable");
            SysData.DataTable clrTable = new System.Data.DataTable("colortable");

            if (lResults.Count == 0)
                return (null, null);

            int maxcol = UseCategoryData() ? lResults[0].Categories.Count : lResults[0].Regulons.Count;


            if (UseCategoryData())
            {
                // count max number of columns neccesary
                for (int r = 1; r < lResults.Count; r++)
                    if (maxcol < lResults[r].Categories.Count)
                        maxcol = lResults[r].Categories.Count;
            }

            else
            {
                // count max number of columns neccesary
                for (int r = 1; r < lResults.Count; r++)
                    if (maxcol < lResults[r].Regulons.Count)
                        maxcol = lResults[r].Regulons.Count;
            }


            // add BSU/gene/p-value/fc columns

            SysData.DataColumn bsuCol = new SysData.DataColumn("bsu", Type.GetType("System.String"));
            myTable.Columns.Add(bsuCol);
            SysData.DataColumn geneCol = new SysData.DataColumn("gene", Type.GetType("System.String"));
            myTable.Columns.Add(geneCol);
            SysData.DataColumn fcCol = new SysData.DataColumn("fc", Type.GetType("System.Double"));
            myTable.Columns.Add(fcCol);
            SysData.DataColumn pValCol = new SysData.DataColumn("pval", Type.GetType("System.Double"));
            myTable.Columns.Add(pValCol);

            SysData.DataColumn funcCol = new SysData.DataColumn("function", Type.GetType("System.String"));
            myTable.Columns.Add(funcCol);
            SysData.DataColumn descCol = new SysData.DataColumn("description", Type.GetType("System.String"));
            myTable.Columns.Add(descCol);

            // add count column
            SysData.DataColumn countCol = new SysData.DataColumn("count_col", Type.GetType("System.Int16"));
            myTable.Columns.Add(countCol);

            // add variable columns
            for (int c = 0; c < maxcol; c++)
            {
                SysData.DataColumn newCol = new SysData.DataColumn(string.Format("col_{0}", c + 1));
                myTable.Columns.Add(newCol);

                if (gSettings.useCat) // add category id columns 
                {
                    SysData.DataColumn catIDCol = new SysData.DataColumn(string.Format("cat_id_{0}", c + 1));
                    myTable.Columns.Add(catIDCol);
                }

                SysData.DataColumn clrCol = new SysData.DataColumn(string.Format("col_{0}", c + 1), Type.GetType("System.Int16"));
                clrTable.Columns.Add(clrCol);
            }

            // fill data from here
            for (int r = 0; r < lResults.Count; r++)
            {
                SysData.DataRow newRow = myTable.Rows.Add();

                newRow["bsu"] = lResults[r].BSU;
                newRow["gene"] = lResults[r].GeneName;
                newRow["fc"] = lResults[r].FC;
                newRow["pval"] = lResults[r].PVALUE;
                newRow["function"] = lResults[r].GeneFunction;
                newRow["description"] = lResults[r].GeneDescription;

                newRow["count_col"] = UseCategoryData() ? lResults[r].Categories.Count : lResults[r].REGULON_TOT;
                SysData.DataRow clrRow = clrTable.Rows.Add();

                if (UseCategoryData())
                {
                    for (int c = 0; c < lResults[r].Categories.Count; c++)
                    {
                        newRow[string.Format("col_{0}", c + 1)] = lResults[r].Categories[c].Name;
                        newRow[string.Format("cat_id_{0}", c + 1)] = lResults[r].Categories[c].catID;
                    }

                    //for (int c = 0; c < lResults[r].REGULON_UP.Count; c++)
                    //    clrRow[lResults[r].REGULON_UP[c]] = 1;

                    //for (int c = 0; c < lResults[r].REGULON_DOWN.Count; c++)
                    //    clrRow[lResults[r].REGULON_DOWN[c]] = -1;
                }
                else
                {
                    for (int c = 0; c < lResults[r].Regulons.Count; c++)
                        newRow[string.Format("col_{0}", c + 1)] = lResults[r].Regulons[c].Name;

                    //for (int c = 0; c < lResults[r].REGULON_UP.Count; c++)
                    //    clrRow[lResults[r].REGULON_UP[c]] = 1;

                    //for (int c = 0; c < lResults[r].REGULON_DOWN.Count; c++)
                    //    clrRow[lResults[r].REGULON_DOWN[c]] = -1;
                }

            }

            return (myTable, clrTable);


        }


        /// <summary>
        /// Reformat the 'raw' augmented data into a datatable that can be displayed on a worksheet
        /// </summary>
        /// <param name="aList"></param>
        /// <returns></returns>
        private SysData.DataTable ReformatRegulonResults(List<FC_BSU> aList)
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
        /// Create an overview of the data mapped to their operons
        /// </summary>
        /// <param name="table"></param>
        private void CreateOperonSheet(SysData.DataTable table)
        {
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();

            int suffix = FindSheetNames(new string[] { "Mapped" });
            lNewSheet.Name = string.Format("Mapped_{0}", suffix);


            int maxNrGenes = Int32.Parse(table.Compute("max([nrgenes])", string.Empty).ToString());

            int infoColumns = 10;

            gApplication.ScreenUpdating = false;
            gApplication.DisplayAlerts = false;
            gApplication.EnableEvents = false;

            int col = 1;
            lNewSheet.Cells[1, col++] = "BSU";
            lNewSheet.Cells[1, col++] = "FC";
            lNewSheet.Cells[1, col++] = "P-Value";
            lNewSheet.Cells[1, col++] = "Gene";
            lNewSheet.Cells[1, col++] = "Gene Function";
            lNewSheet.Cells[1, col++] = "Gene Description";
            lNewSheet.Cells[1, col++] = "Operon Name";
            lNewSheet.Cells[1, col++] = "Nr operons";
            lNewSheet.Cells[1, col++] = "Nr genes";
            lNewSheet.Cells[1, col++] = "Operon";

            for (int c = 0; c < maxNrGenes; c++)
            {
                string colHeader = string.Format("FC Gene #{0}", c + 1);
                lNewSheet.Cells[1, c + col] = colHeader;
            }

            FastDtToExcel(table, lNewSheet, 2, 1, table.Rows.Count + 1, maxNrGenes + infoColumns);


            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[table.Rows.Count + 1, maxNrGenes + infoColumns];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            all.Columns.AutoFit();
            all.Rows.AutoFit();


            gApplication.ScreenUpdating = true;
            gApplication.DisplayAlerts = true;
            gApplication.EnableEvents = true;
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
        /// Create a worksheet that contains the extended (=all genes) data per category/regulon. Used in spreading plot
        /// </summary>
        /// <param name="theElements"></param>
        /// <param name="chartName"></param>
        private void CreateExtendedRegulonCategoryDataSheet(element_fc theElements, string chartName)
        {

            string sheetName = chartName.Replace("Plot", "Tab");
            //aSheet.Name = sheetName;

            string catRegLabel = Properties.Settings.Default.useCat ? "Category" : "Regulon";

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

        private (Excel.Worksheet, List<summaryInfo>) CreateRankingDataSheet(element_fc theElements, List<summaryInfo> all, List<summaryInfo> posSort, List<summaryInfo> negSort, List<summaryInfo> bestSort, int suffix, bool detailSheet = false)
        {
            string catRegLabel = Properties.Settings.Default.useCat ? "CatRankTable" : "RegRankTable";
            if (detailSheet)
                catRegLabel = "Mapping_Details";
            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();

            catRegLabel = String.Format("{0}_{1}", catRegLabel, suffix);
            lNewSheet.Name = catRegLabel;

            //RenameWorksheet(lNewSheet, catRegLabel);

            DataTable lTable = ElementsToTable(all);

            string catRegHeader = Properties.Settings.Default.useCat ? "Category" : "Regulon";
            string firstBlockHeader = Properties.Settings.Default.useCat ? "Plot data" : "Without regulatory directionality";
            string secondBlockHeader = Properties.Settings.Default.useCat ? "Positive fc" : "When regulator is activated";
            string thirdBlockHeader = Properties.Settings.Default.useCat ? "Negative fc" : "When regulator is repressed";
            string fourthBlockHeader = Properties.Settings.Default.useCat ? "Best results" : "Best score";
            string FCheader = Properties.Settings.Default.useCat ? "Average FC" : "Average FC";
            string MADheader = Properties.Settings.Default.useCat ? "MAD FC" : "MAD FC";

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
            lNewSheet.Cells[hdrRow, 23] = "Average FC";
            lNewSheet.Cells[hdrRow, 24] = "MAD FC";
            lNewSheet.Cells[hdrRow, 25] = "Average P";

            // Combine positive and negative mode results to obtain a 'best' result

            lTable = ElementsToTable(bestSort, bestMode: true);

            //List<summaryInfo> _best;
            //    (lTable, _best) = BestElementScore(theElements);

            FastDtToExcel(lTable, lNewSheet, hdrRow + 1, 19, lTable.Rows.Count + hdrRow, lTable.Columns.Count + 18);

            top = lNewSheet.Cells[1, 19];
            bottom = lNewSheet.Cells[lTable.Rows.Count + hdrRow, 25];
            rall = (Excel.Range)lNewSheet.get_Range(top, bottom);
            rall.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
            rall.Interior.TintAndShade = 0.8;
            rall.Interior.PatternTintAndShade = 0;

            return (lNewSheet, null);


        }

        private void AdjustColumns(Excel.Range aRange)
        {
            aRange.ColumnWidth = 17.0;
        }


        /// <summary>
        /// Create the worksheet that contains the basic mapping gene - regulon table
        /// </summary>
        /// <param name="bsuRegulons"></param>
        private int CreateMappingSheet(SysData.DataTable aTable, List<summaryInfo> bestInfo)
        {

            DataTable workTable = aTable.Copy();

            if (gSettings.useCat) // remove category id columns in case of regulon output
            {
                object _mc = workTable.Compute("Max(count_col)", "");
                int maxCount = Int32.Parse(_mc.ToString());
                for (int c = 0; c < maxCount; c++)
                {
                    string _colFmt = String.Format("cat_id_{0}", c + 1);
                    workTable.Columns.Remove(_colFmt);
                }

            }

            if (bestInfo != null)
            {

                for (int r = 0; r < workTable.Rows.Count; r++)
                {
                    DataRow dataRow = workTable.Rows[r];
                    int countCol = Int32.Parse(dataRow["count_col"].ToString());
                    for (int c = 0; c < countCol; c++)
                    {
                        string colFmt = String.Format("col_{0}", c + 1);
                        string _val = dataRow[colFmt].ToString();

                        summaryInfo lItem = bestInfo.Find(item => item.catName == _val);
                        string strAvgFC = lItem.fc_average.ToString("0.00", CultureInfo.InvariantCulture);
                        strAvgFC = strAvgFC.Replace(',', '.');
                        _val = String.Format("{0}(FC:{1} {2:0}%)", _val, strAvgFC, lItem.best_gene_percentage);

                        dataRow.BeginEdit();
                        dataRow[colFmt] = _val;
                        dataRow.EndEdit();


                    }
                }
            }

            AddTask(TASKS.UPDATE_MAPPED_TABLE);

            int nrRows = workTable.Rows.Count;
            int startR = 2;
            int offsetColumn = 1;

            Excel.Worksheet lNewSheet = gApplication.Worksheets.Add();
            int suffix = FindSheetNames(new string[] { "Mapped", "Mapped_Details" });
            lNewSheet.Name = String.Format("Mapped_{0}", suffix);

            lNewSheet.Cells[1, 1] = "Gene id";
            lNewSheet.Cells[1, 2] = "Gene name";
            lNewSheet.Cells[1, 3] = "Fold-change";
            lNewSheet.Cells[1, 4] = "P-value";
            lNewSheet.Cells[1, 5] = "Function";
            lNewSheet.Cells[1, 6] = "Description";
            lNewSheet.Cells[1, 7] = UseCategoryData() ? "Tot# Categories" : "Tot# Regulons";

            string lastColumn = workTable.Columns[workTable.Columns.Count - 1].ColumnName;
            lastColumn = lastColumn.Replace("col_", "");
            int maxreg = ClassExtensions.ParseInt(lastColumn, 0);

            for (int i = 0; i < maxreg; i++)
                lNewSheet.Cells[1, i + 8] = string.Format(UseCategoryData() ? "Category_{0}" : "Regulon_{0}", i + 1);

            // copy data to excel sheet

            FastDtToExcel(workTable, lNewSheet, startR, offsetColumn, startR + nrRows - 1, offsetColumn + workTable.Columns.Count - 1);

            Excel.Range top = lNewSheet.Cells[1, 1];
            Excel.Range bottom = lNewSheet.Cells[workTable.Rows.Count + 1, workTable.Columns.Count];
            Excel.Range all = (Excel.Range)lNewSheet.get_Range(top, bottom);

            AdjustColumns(all);
            //all.Columns.AutoFit();           

            // color cells according to table 
            //ColorCells(clrTbl, lNewSheet, startR, offsetColumn + 5, startR + nrRows - 1, offsetColumn + lTable.Columns.Count - 1);

            RemoveTask(TASKS.UPDATE_MAPPED_TABLE);

            return suffix;

        }


    }
}
