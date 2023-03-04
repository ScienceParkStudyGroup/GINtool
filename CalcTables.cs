using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;
using static GINtool.ES_Functions;
using stat_dict = System.Collections.Generic.Dictionary<string, double>;
using rank_dict = System.Collections.Generic.Dictionary<string, int>;
using dict_rank = System.Collections.Generic.Dictionary<int, string>;
using lib_dict = System.Collections.Generic.Dictionary<string, string[]>;
using Microsoft.Office.Core;


namespace GINtool
{
    public partial class GinRibbon
    {
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

        ///// <summary>
        ///// Create a basic count / average usage table per regulon based on the regulon data from subtwiki
        ///// Because regulator_BSU is not known for all regulators whe use the regulator name
        ///// </summary>
        //private void CreateTableStatistics()
        //{
        //    List<string> lString = new List<string> { Properties.Settings.Default.referenceRegulon };
        //    SysData.DataTable lUnique = GetDistinctRecords(gRegulonWB, lString.ToArray());

        //    // initialize the global datatable

        //    gRefStats = new SysData.DataTable("tblstat");

        //    int totNrRows = gRegulonWB.Rows.Count;

        //    SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
        //    SysData.DataColumn countColumn = new SysData.DataColumn("Count", Type.GetType("System.Int16"));
        //    SysData.DataColumn avgColumn = new SysData.DataColumn("Average", Type.GetType("System.Double"));
        //    gRefStats.Columns.Add(regColumn);
        //    gRefStats.Columns.Add(countColumn);
        //    gRefStats.Columns.Add(avgColumn);

        //    foreach (SysData.DataRow lRow in lUnique.Rows)
        //    {
        //        string lVal = lRow[Properties.Settings.Default.referenceRegulon].ToString();
        //        int cnt = gRegulonWB.Select(string.Format("{0}='{1}'", Properties.Settings.Default.referenceRegulon, lVal)).Length;
        //        SysData.DataRow nRow = gRefStats.Rows.Add();
        //        nRow["Regulon"] = lVal;
        //        nRow["Count"] = cnt;
        //        nRow["Average"] = ((double)cnt) / totNrRows;
        //    }
        //}

        private SysData.DataTable MappedCategoryTable()
        {
            return null;
        }

        private SysData.DataTable MappedOperonTable()
        {
            return null;
        }

        /// <summary>
        /// The main routine after mouse selection update. It reads in the data from the excel sheet to return Raw input in a list of BSU structures and a List of genes and their associated regulons 
        /// </summary>
        /// <param name="suppressOutput"></param>
        /// <returns></returns>        
        private List<BsuLinkedItems> ReadDataAndAugment()
        {

            AddTask(TASKS.READ_SHEET_DATA);

            List<Excel.Range> theInputCells = new List<Excel.Range>()
            {
                gRangeP,
                gRangeFC,
                gRangeBSU
            };

            // set flag is data has changed       
            //if (InputHasChanged() || gOutput == null || gList == null)
            if (InputHasChanged() || gList == null)
            {
                gNeedsUpdate = (byte)UPDATE_FLAGS.ALL;
            }
            else
            {
                RemoveTask(TASKS.READ_SHEET_DATA);
                return gList;
            }

            int nrRows = gRangeP.Rows.Count;

            // generate the results for outputting the data and summary
            try
            {
                List<BsuLinkedItems> lResults = AugmentWithGeneInfo(theInputCells);

                if (CanAugmentWithCategoryData())
                    lResults = AugmentWithCategoryData(lResults);
                if (CanAugmentWithRegulonData())
                    lResults = AugmentWithRegulonData(lResults);

                RemoveTask(TASKS.READ_SHEET_DATA);

                CalibrateES();

                return lResults;
            }
            catch
            {
                MessageBox.Show("Are you sure the columns are properly mapped?");
                RemoveTask(TASKS.READ_SHEET_DATA);

                return null;
            }

        }

        private void CalibrateES()
        {

            if (gCategoryDict.Count == 0 && gRegulonDict.Count == 0)
                return;

            AddTask(TASKS.ES_CALIBRATION);

            // gsea_calibrate(stat_dict signature, lib_dict library, ref Hashtable hashtable, int permutations = 2000, int anchors = 20, int min_size = 5, int max_size = 10000, bool verbose = false, bool symmetric = true, bool signature_cache = true, bool shared_null = false, int seed = 0)
            stat_dict _tmp = new stat_dict();
            foreach (KeyValuePair<string, DataItem> record in gDataSetDict)
            {
                _tmp.Add(record.Key, record.Value.FC);
            }
            gsea_calibrate(_tmp, gRegulonDict, ref gFgseaHash, min_size:1);
            
            RemoveTask(TASKS.ES_CALIBRATION);

            AddTask(TASKS.ES_CALCULATION);
            gGSEA_dict = gsea_enrich(_tmp, gRegulonDict, gFgseaHash, min_size: 1);
            RemoveTask(TASKS.ES_CALCULATION);
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
        ///  Create a combined table that combines the raw data and regulon summaries
        /// </summary>
        /// <param name="aUsageTbl">the tabel with mapped categories/regulons</param>
        /// <param name="lLst"></param>
        /// <returns></returns>
        private (SysData.DataTable, SysData.DataTable) CreateCombinedTable(SysData.DataTable aUsageTbl, List<BsuLinkedItems> lLst, List<summaryInfo> bestResults)
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
                if (maxRegulons < lLst[i].Regulons.Count)
                    maxRegulons = lLst[i].Regulons.Count;
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

                if (accept)
                {
                    SysData.DataRow lColorRow = lColorTable.Rows.Add();

                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["FC"] = lLst[r].FC;
                    lRow["BSU"] = lLst[r].BSU;
                    lRow["GENE"] = lLst[r].GeneName;
                    lRow["PVALUE"] = lLst[r].PVALUE;


                    double FC = lLst[r].FC;

                    for (int i = 0; i < lLst[r].Regulons.Count; i++)
                    {

                        // check association direction 
                        bool posAssoc = lLst[r].REGULON_UP.Contains(i);
                        bool negAssoc = lLst[r].REGULON_DOWN.Contains(i);
                        // depending on the association in the table the cell color is red or green

                        int clrInt = posAssoc ? 1 : negAssoc ? -1 : 0;

                        SysData.DataRow[] lHit = aUsageTbl.Select(string.Format("Regulon = '{0}'", lLst[r].Regulons[i].Name));
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

                        lRow[string.Format("Regulon_{0}", i + 1)] = lLst[r].Regulons[i] + " " + lVal;
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
        /// Create the data table that can be used to display the data augmented with operon info in an Excel sheet.
        /// </summary>
        /// <param name="aUsageTbl"></param>
        /// <param name="lLst"></param>
        /// <returns></returns>
        private SysData.DataTable CreateOperonTable(List<BsuLinkedItems> lLst)
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


            col = new SysData.DataColumn("Function", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("Description", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("operon", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("nroperons", Type.GetType("System.Int16"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("nrgenes", Type.GetType("System.Int16"));
            lTable.Columns.Add(col);

            col = new SysData.DataColumn("operon_genes", Type.GetType("System.String"));
            lTable.Columns.Add(col);

            for (int nr = 0; nr < gMaxGenesPerOperon; nr++) // remove extra columns later if necessary
            {
                col = new SysData.DataColumn(string.Format("gene_{0}", nr + 1), Type.GetType("System.Double"));
                lTable.Columns.Add(col);
            }

            int maxColumnsUsed = 0;


            //List<string> genesIDs = new List<string>();
            //foreach(BsuLinkedItems item in lLst)            
            //    genesIDs.Add("'"+item.GeneName+"'");

            //string _filter = String.Join(",", genesIDs.ToArray());
            //SysData.DataRow[] _lt = gRefOperonsWB.Select(String.Format("gene in ({0})", _filter));

            //double lowVal = Properties.Settings.Default.fcLOW;

            // loop over the all the genes found
            for (int r = 0; r < lLst.Count; r++)
            {

                string geneName = lLst[r].GeneName;
                //double lFC = lLst[r].FC;
                //double lPval = lLst[r].PVALUE;

                List<string> luOperons = new List<string>();


                // possibly multiple operons for a single gene
                SysData.DataRow[] lOperons = gRefOperonsWB.Select(string.Format("gene='{0}'", geneName));


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


                // add a row
                SysData.DataRow lRow = lTable.Rows.Add();
                lRow["BSU"] = lLst[r].BSU;
                lRow["FC"] = lLst[r].FC;
                lRow["P-Value"] = lLst[r].PVALUE;
                lRow["Function"] = lLst[r].GeneFunction;
                lRow["Description"] = lLst[r].GeneDescription;
                lRow["gene"] = geneName;

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

                    lRow["operon"] = operon;
                    lRow["nroperons"] = noperons;
                    lRow["nrgenes"] = nrgenes;
                    lRow["operon_genes"] = opgenes;



                    if (nrgenes > maxColumnsUsed)
                        maxColumnsUsed = nrgenes;

                    // copy the FCs

                    for (int i = 0; i < nrgenes; i++)
                    {
                        if (!(lFCs[i] is Double.NaN))
                            lRow[string.Format("gene_{0}", i + 1)] = lFCs[i];
                    }
                }

            }

            for (int _g = gMaxGenesPerOperon; _g > maxColumnsUsed; _g--)
            {
                string colFmt = String.Format("gene_{0}", _g);
                lTable.Columns.Remove(colFmt);
            }
            return lTable;
        }

        /// <summary>
        /// Create the linkage table between regulon and genes
        /// </summary>
        /// <param name="aList"></param>
        /// <returns></returns>
        private SysData.DataTable CreateGeneUsageTable(List<BsuLinkedItems> aList)
        {

            SysData.DataTable lTable = new SysData.DataTable("GeneResultsTable");
            SysData.DataColumn geneColumn = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            SysData.DataColumn geneIDcolumn = new SysData.DataColumn("Gene_ID", Type.GetType("System.String"));

            /* Need to check if we do need all of them */
            SysData.DataColumn pvalColumn = new SysData.DataColumn("Pvalue", Type.GetType("System.Double"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            SysData.DataColumn functionColumn = new SysData.DataColumn("GeneFunction", Type.GetType("System.String"));
            SysData.DataColumn decsColumn = new SysData.DataColumn("GeneDescription", Type.GetType("System.String"));



            //lTable.Columns.Add(regColumn);
            lTable.Columns.Add(geneIDcolumn);
            lTable.Columns.Add(geneColumn);
            lTable.Columns.Add(fcColumn);
            lTable.Columns.Add(pvalColumn);
            //lTable.Columns.Add(dirColumn);
            lTable.Columns.Add(functionColumn);
            lTable.Columns.Add(decsColumn);

            foreach (BsuLinkedItems _it in aList)
            {
                SysData.DataRow lRow = lTable.Rows.Add();
                lRow["Gene"] = _it.GeneName;
                lRow["Gene_ID"] = _it.BSU;
                lRow["FC"] = _it.FC;
                lRow["GeneFunction"] = _it.GeneFunction;
                lRow["GeneDescription"] = _it.GeneDescription;
                lRow["Pvalue"] = _it.PVALUE;
            }

            return lTable;
        }


        /// <summary>
        /// Create the linkage table between regulon and genes
        /// </summary>
        /// <param name="aList"></param>
        /// <returns></returns>
        private SysData.DataTable CreateRegulonUsageTable(List<BsuLinkedItems> aList)
        {

            SysData.DataTable lTable = new SysData.DataTable("MappedRegulons");
            SysData.DataColumn regColumn = new SysData.DataColumn("Regulon", Type.GetType("System.String"));
            SysData.DataColumn geneColumn = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            SysData.DataColumn geneIDcolumn = new SysData.DataColumn("Gene_ID", Type.GetType("System.String"));

            /* Need to check if we do need all of them */
            SysData.DataColumn pvalColumn = new SysData.DataColumn("Pvalue", Type.GetType("System.Double"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            SysData.DataColumn dirColumn = new SysData.DataColumn("DIR", Type.GetType("System.Int32"));
            SysData.DataColumn functionColumn = new SysData.DataColumn("GeneFunction", Type.GetType("System.String"));
            SysData.DataColumn decsColumn = new SysData.DataColumn("GeneDescription", Type.GetType("System.String"));



            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(geneIDcolumn);
            lTable.Columns.Add(geneColumn);
            lTable.Columns.Add(fcColumn);
            lTable.Columns.Add(pvalColumn);
            lTable.Columns.Add(dirColumn);
            lTable.Columns.Add(functionColumn);
            lTable.Columns.Add(decsColumn);

            foreach (BsuLinkedItems _it in aList)
            {
                List<RegulonItem> _regulons = _it.Regulons;
                foreach (RegulonItem __it in _regulons)
                {
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["Regulon"] = __it.Name;
                    lRow["DIR"] = __it.Direction == "UP" ? 1 : ( __it.Direction == "DOWN" ? -1 : 0); // add check for when no direction is given
                    lRow["Gene_ID"] = _it.BSU;
                    lRow["Gene"] = _it.GeneName;
                    lRow["FC"] = _it.FC;
                    lRow["GeneFunction"] = _it.GeneFunction;
                    lRow["GeneDescription"] = _it.GeneDescription;
                    lRow["Pvalue"] = _it.PVALUE;
                }

            }


            return lTable;
        }

        /// <summary>
        /// Create the linkage tables between categories and genes
        /// </summary>
        /// <param name="aList"></param>
        /// <returns></returns>
        private SysData.DataTable CreateCategoryUsageTable(List<BsuLinkedItems> aList)
        {

            SysData.DataTable lTable = new SysData.DataTable("MappedCategories");
            SysData.DataColumn regColumn = new SysData.DataColumn("Category", Type.GetType("System.String"));
            SysData.DataColumn geneColumn = new SysData.DataColumn("Gene", Type.GetType("System.String"));
            SysData.DataColumn geneIDcolumn = new SysData.DataColumn("Gene_ID", Type.GetType("System.String"));
            /* Need to check if we do need all of them */
            SysData.DataColumn pvalColumn = new SysData.DataColumn("Pvalue", Type.GetType("System.Double"));
            SysData.DataColumn fcColumn = new SysData.DataColumn("FC", Type.GetType("System.Double"));
            //SysData.DataColumn dirColumn = new SysData.DataColumn("DIR", Type.GetType("System.Int32"));
            SysData.DataColumn functionColumn = new SysData.DataColumn("GeneFunction", Type.GetType("System.String"));
            SysData.DataColumn decsColumn = new SysData.DataColumn("GeneDescription", Type.GetType("System.String"));



            lTable.Columns.Add(regColumn);
            lTable.Columns.Add(geneIDcolumn);
            lTable.Columns.Add(geneColumn);
            lTable.Columns.Add(fcColumn);
            lTable.Columns.Add(pvalColumn);
            //lTable.Columns.Add(dirColumn);
            lTable.Columns.Add(functionColumn);
            lTable.Columns.Add(decsColumn);

            foreach (BsuLinkedItems _it in aList)
            {
                List<CategoryItem> _categories = _it.Categories;
                foreach (CategoryItem __it in _categories)
                {
                    SysData.DataRow lRow = lTable.Rows.Add();
                    lRow["Category"] = __it.Name;
                    lRow["Gene_ID"] = _it.BSU;
                    lRow["Gene"] = _it.GeneName;
                    lRow["FC"] = _it.FC;
                    lRow["GeneFunction"] = _it.GeneFunction;
                    lRow["GeneDescription"] = _it.GeneDescription;
                    lRow["Pvalue"] = _it.PVALUE;
                }

            }


            return lTable;
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


        private void CreateBestDataTable(List<BsuLinkedItems> bsuRegulons)
        {

            (SysData.DataTable lMappingTable, SysData.DataTable clrTbl) = PrepareResultTable(bsuRegulons);


            if (lMappingTable is null)
                return;


            List<cat_elements> cat_Elements = new List<cat_elements>();

            string lastColumn = lMappingTable.Columns[lMappingTable.Columns.Count - 1].ColumnName;
            lastColumn = lastColumn.Replace("col_", "");
            //int maxreg = ClassExtensions.ParseInt(lastColumn, 0);

            for (int row = 0; row < lMappingTable.Rows.Count; row++)
            {

                DataRow dataRow = lMappingTable.Rows[row];
                int nrcol = Int32.Parse(dataRow["count_col"].ToString());
                string _bsu = dataRow["bsu"].ToString();

                for (int col = 0; col < nrcol; col++)
                {
                    string colFmt = String.Format("col_{0}", col + 1);
                    string _catName = dataRow[colFmt].ToString();

                    string _catID = "";

                    if (gSettings.useCat)
                    {
                        colFmt = String.Format("cat_id_{0}", col + 1);
                        _catID = dataRow[colFmt].ToString();
                    }

                    cat_elements cat_Elements2 = new cat_elements();

                    cat_Elements2.catName = _catName;
                    cat_Elements2.elTag = _catID;
                    cat_Elements2.elements = new string[] { _catID };
                    cat_Elements.Add(cat_Elements2);
                }
            }


            if (gSettings.useOperons)
            {
                SysData.DataTable tblOperon = CreateOperonTable(bsuRegulons);
                CreateOperonSheet(tblOperon);
            }
            else
            {
                cat_Elements = GetUniqueElements(cat_Elements);

                SysData.DataView dataView = gSettings.useCat ? gCategoryTable.AsDataView() : gRegulonTable.AsDataView();
                element_fc catPlotData;
                if (gSettings.useCat)
                    catPlotData = CatElements2ElementsFC(dataView, cat_Elements);
                else
                    catPlotData = Regulons2ElementsFC(dataView, cat_Elements);


                if (catPlotData.All != null)
                {
                    (List<element_rank> plotData, List<summaryInfo> _all, List<summaryInfo> _pos, List<summaryInfo> _neg, List<summaryInfo> _best) = CreateRankingPlotData(catPlotData);

                    int suffix = CreateMappingSheet(lMappingTable, _best);
                    //if (outputDetails)
                    CreateRankingDataSheet(catPlotData, _all, _pos, _neg, _best, suffix, detailSheet: true);
                }
            }
        }

        /// <summary>
        /// Order the ranking results by name and create bubble plot data
        /// </summary>
        /// <param name="theElements"></param>
        /// <returns></returns>
        private (List<element_rank>, List<summaryInfo>, List<summaryInfo>, List<summaryInfo>, List<summaryInfo>) CreateRankingPlotData(element_fc theElements)
        {

            List<summaryInfo> all_elements = SortedElements(theElements.All, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> pos_elements = SortedElements(theElements.Activated, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> neg_elements = SortedElements(theElements.Repressed, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> best_elements = null;
            (gBestTable, best_elements) = BestElementScore(theElements);

            return (BubblePlotData(theElements.All), all_elements, pos_elements, neg_elements, best_elements);


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
                summaryInfo _all = element_info.All.GetCatValues(catName); // update 18_10_22
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
                
                if ( (n1 == 0) & (n2 == 0))
                {
                    lRow["Mode"] = "not defined";                    
                }

                lRow["Count"] = _si1.genes.Length;
                if (totnrgenes > 0)
                    lRow["Percentage"] = Math.Round((double)_si1.genes.Length / (double)totnrgenes * 100);

                _si1.best_gene_percentage = Math.Round((double)_si1.genes.Length / (double)totnrgenes * 100);

                 
                if(( n1 ==0 ) & (n2 == 0)) // if no direction was specicified copy the data from the _all structure
                {
                    _all.best_gene_percentage = 0.0;
                    _output.Add(_all);
                }
                else // otherwise copy the best one
                {
                    _output.Add(_si1);
                }

                

                if (n1 > 0) // _s1 always contains the 'best' direction.
                {
                    lRow["AverageABSFC"] = _si1.fc_average;
                    lRow["MadABSFC"] = _si1.fc_mad;
                    lRow["AverageP"] = _si1.p_average;
                }
                else // n1 == 0 & n2 == 0, added 19-10-22
                {
                    lRow["AverageABSFC"] = _all.fc_average;
                    lRow["MadABSFC"] = _all.fc_mad;
                    lRow["AverageP"] = _all.p_average;
                    lRow["Percentage"] = 0.0;
                }

            }

            DataView _dv = lTable.DefaultView;
            _dv.Sort = "Name asc";

            return (_dv.ToTable(), _output);
        }


    }

}
