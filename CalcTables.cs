﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SysData = System.Data;
using static GINtool.ES_Functions;
using Accord.Statistics.Distributions.Univariate;
using static GINtool.ES_Extensions;
using Accord.Math;
using System.Collections;
using System.Windows.Input;
using Accord.Math.Distances;
using System.Windows.Markup;
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
            if (InputHasChanged() || gList == null || gList.Count==0)
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
                //gFgseaHash.Clear();
                //gDataSetStat_dict.Clear();
                //gGSEADict.Clear();
                gBSU_gene_dict.Clear();

                List<BsuLinkedItems> lResults = AugmentWithGeneInfo(theInputCells);

                if (CanAugmentWithCategoryData())
                    lResults = AugmentWithCategoryData(lResults);
                if (CanAugmentWithRegulonData())
                    lResults = AugmentWithRegulonData(lResults);

                RemoveTask(TASKS.READ_SHEET_DATA);

                // make sure that there are no duplicate keys ... not expected 
                gCombinedDict = gRegulonDict.Concat(gCategoryDict).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                
                
                // CheckValues();

                CalibrateES();

                return lResults;
            }
            catch (Exception ex) 
            {
                MessageBox.Show("Are you sure the columns are properly mapped? " + ex.Message);
                RemoveTask(TASKS.READ_SHEET_DATA);

                return null;
            }

        }

        private void CheckValues()
        {
            Dictionary<string, int> hashValues = new Dictionary<string, int>();
            Dictionary<string, double> signature = getSignature(gDataSetDict, !Properties.Settings.Default.gseaFC); 
            // gDataSetDict.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => Math.Sign(kvp.Value.FC) * -Math.Log10(kvp.Value.pval));

            //Dictionary<string, double> signature = gDataSetDict.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => Math.Abs(kvp.Value.FC));

            signature = signature.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
            //dict_rank map_signature = signature.MapRank();
            //rank_dict signature_map = signature.RankMap();

            //NormalDistribution norm = new NormalDistribution();

            string[] signature_genes = signature.Keys.ToArray();

            int min_size = 1, max_size = 25000;

            foreach (string key in gCombinedDict.Keys)
            {
                string[] gene_set = gCombinedDict[key];
                string[] stripped_set = strip_gene_set(signature_genes, gene_set);
                if (stripped_set.Length >= min_size && stripped_set.Length <= max_size)
                {
                    int gsHash = stripped_set.GetHashCodeValue();
                    hashValues.Add(key, gsHash);
                }
                string[] gene_set_pos = stripped_set.Where(k => gDataSetDict[k].FC > 0).ToArray();
                string[] stripped_set_pos = strip_gene_set(signature_genes, gene_set_pos);
                if (stripped_set_pos.Length >= min_size && stripped_set_pos.Length <= max_size)
                {
                    int gsHash = stripped_set_pos.GetHashCodeValue();
                    hashValues.Add(String.Format("{0}_pos", key), gsHash);

                }
                string[] gene_set_neg = stripped_set.Where(k => gDataSetDict[k].FC < 0).ToArray();
                string[] stripped_set_neg = strip_gene_set(signature_genes, gene_set_neg);

                if (stripped_set_neg.Length >= min_size && stripped_set_neg.Length <= max_size)
                {
                    int gsHash = stripped_set_neg.GetHashCodeValue();
                    hashValues.Add(String.Format("{0}_neg", key), gsHash);
                }
            }

            var distinctList = hashValues.Values.Distinct().ToList();

            Dictionary<int, IEnumerable<string>> myDict = new Dictionary<int, IEnumerable<string>>();
            foreach(int v in  distinctList)
            {
                IEnumerable<string> myKeys = hashValues.Where(pair => pair.Value == v).Select(pair => pair.Key);
                myDict.Add(v, myKeys);
            }

            //System.IO.File.WriteAllLines(@"c:\temp\pathtocsv.csv", myDict.Select(x => x.Key + "," + String.Join(",",x.Value.ToArray()) + ","));


        }


        private void CalibrateES()
        {

            if (gCategoryDict.Count == 0 && gRegulonDict.Count == 0)
                return;

            AddTask(TASKS.ES_CALIBRATION);                       
            gsea_calibrate(gDataSetDict, gCombinedDict, ref gFgseaHash,pvalues:!Properties.Settings.Default.gseaFC);            
            RemoveTask(TASKS.ES_CALIBRATION);

            AddTask(TASKS.ES_CALCULATION);
            gsea_enrich(gDataSetDict, gCombinedDict, gFgseaHash, ref gGSEAHash, min_size: 1, pvalues: !Properties.Settings.Default.gseaFC);                        
            RemoveTask(TASKS.ES_CALCULATION);

            CopyStaticESParameters();


        }

        private void CopyStaticESParameters()
        {
            gES_signature = getSignature(gDataSetDict,!Properties.Settings.Default.gseaFC);

            gES_signature_ordered = gES_signature.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
            gES_map_signature = gES_signature_ordered.MapRank();
            gES_signature_map = gES_signature_ordered.RankMap();
            NormalDistribution normal = new NormalDistribution();
            gES_signature_genes = gES_signature_ordered.Keys.ToArray();                        
            gES_sigvalues = gES_signature_ordered.Values.ToArray();
            gES_sigvalues = gES_sigvalues.Plus(Pmult(normal.Generate(gES_sigvalues.Length), 1 / (gES_sigvalues.Average() * 10000)));
            gES_abs_signature = gES_sigvalues.Abs();

            Dictionary<string,double> signature_ordered = gES_signature.OrderBy(kvp => kvp.Value).ToDictionary(x => x.Key, x => x.Value);
            int sighashK = signature_ordered.Keys.GetHashCodeValue();
            int sighashV = signature_ordered.Values.GetHashCodeValue();
            int sig_hash = sighashK + sighashV;
            gES_key = sighashK + sighashV;


        }

        private (double,double) CalcES(IEnumerable<string> gene_set)
        {
            S_GSEA result = gsea_calc(gES_abs_signature, gES_signature_genes, gES_map_signature, gES_signature_map, gene_set, (S_ESPARAMS) gFgseaHash[gES_key], ref gGSEAHash, min_size: 1);
            double pval = result.pval == 0 ? 1e-80 : result.pval;
            return (result.es, pval);
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
        private List<element_rank> BubblePlotData(List<summaryInfo> info, bool volcanoPlot=false, int maxExtreme=-1)
        {
            List<element_rank> element_Ranks = new List<element_rank>();

            // MAD values
            List<double> e1_m = new List<double>(), e2_m = new List<double>(), e3_m = new List<double>(), e4_m = new List<double>(), e5_m = new List<double>();
            // FC values
            List<double> e1_fc = new List<double>(), e2_fc = new List<double>(), e3_fc = new List<double>(), e4_fc = new List<double>(), e5_fc = new List<double>();
            // Counts
            List<int> e1_n = new List<int>(), e2_n = new List<int>(), e3_n = new List<int>(), e4_n = new List<int>(), e5_n = new List<int>();
            // p_values
            List<double> e1_pv = new List<double>(), e2_pv = new List<double>(), e3_pv = new List<double>(), e4_pv = new List<double>(), e5_pv = new List<double>();
            // CATEGORY/REGULON NAMES
            List<string> e1_s = new List<string>(), e2_s = new List<string>(), e3_s = new List<string>(), e4_s = new List<string>(), e5_s = new List<string>();
            // BEST GENE PERCENTAGES
            List<double> e1_p = new List<double>(), e2_p = new List<double>(), e3_p = new List<double>(), e4_p = new List<double>(), e5_p = new List<double>();


            double[] pvalues = info.Select(x=>x.p_fdr).OrderBy(x=>x).ToArray();
            double[] q = new double[] { 10D, 20D, 33D, 50D };
            double[] nn = percentiles(pvalues, q).OrderBy(x=>x).ToArray();


            foreach (summaryInfo sInfo in info)
            {

                double pVal = sInfo.p_fdr;

                bool check1 = pVal <= nn[0];
                bool check2 = pVal > nn[0] && pVal <= nn[1];
                bool check3 = pVal > nn[1] && pVal <= nn[2];
                bool check4 = pVal > nn[2] && pVal <= nn[3];
                bool check5 = pVal > nn[3];

                if (volcanoPlot)
                {
                    check1 = sInfo.best_gene_percentage >= 90;
                    check2 = sInfo.best_gene_percentage >= 80 && sInfo.best_gene_percentage < 90;
                    check3 = sInfo.best_gene_percentage >= 75 && sInfo.best_gene_percentage < 80; 
                    check4 = sInfo.best_gene_percentage >= 67 && sInfo.best_gene_percentage < 75;
                    check5 = sInfo.best_gene_percentage >0  && sInfo.best_gene_percentage < 67;
                }               


                List<double> _workfc = null;
                List<double> _workm = null;
                List<int> _workn = null;
                List<string> _works = null;
                List<double> _workp = null;
                List<double> _workpv = null;


                if (check1 && sInfo.genes[0] != "")
                {
                    _workfc = e1_fc;
                    _workm = e1_m;
                    _workn = e1_n;
                    _works = e1_s;
                    _workp = e1_p;
                    _workpv = e1_pv;
                }

                if (check2 && sInfo.genes[0] != "")
                {
                    _workfc = e2_fc;
                    _workm = e2_m;
                    _workn = e2_n;
                    _works = e2_s;
                    _workp = e2_p;
                    _workpv = e2_pv;

                }


                if (check3 && sInfo.genes[0] != "")
                {
                    _workfc = e3_fc;
                    _workm = e3_m;
                    _workn = e3_n;
                    _works = e3_s;
                    _workp = e3_p;
                    _workpv = e3_pv;

                }
                if (check4 && sInfo.genes[0] != "")
                {
                    _workfc = e4_fc;
                    _workm = e4_m;
                    _workn = e4_n;
                    _works = e4_s;
                    _workp = e4_p;
                    _workpv = e4_pv;

                }


                if (check5 && sInfo.genes[0] != "")
                {
                    _workfc = e5_fc;
                    _workm = e5_m;
                    _workn = e5_n;
                    _works = e5_s;
                    _workp = e5_p;
                    _workpv = e5_pv;

                }

                if (_workfc != null)
                {

                    _workfc.Add(sInfo.fc_average);
                    _workm.Add(sInfo.fc_mad);
                    _workn.Add(sInfo.p_values != null ? sInfo.p_values.Length : 0);
                    double _pVal = -Math.Log10(sInfo.p_fdr);
                    _workpv.Add( maxExtreme >0 ? (_pVal > maxExtreme ? maxExtreme : _pVal) : _pVal);
                    _works.Add(StripText(sInfo.catName));
                    _workp.Add(sInfo.best_gene_percentage);
                }

            }
            
            string catLabel1 = "<=10%", catLabel2 = "10%-20%", catLabel3 = "20%-33%", catLabel4 = "33%-50%", catLabel5 = ">50%";
            if (volcanoPlot)
            {
                catLabel1 = ">=90";
                catLabel2 = ">=80";
                catLabel3 = ">=75";
                catLabel4 = ">=67";
                catLabel5 = ">=50";
            }

            element_rank e1 = new element_rank()
            {
                catName = catLabel1,
                p_fdr = e1_pv.ToArray(),
                average_fc = e1_fc.ToArray(),
                mad_fc = e1_m.ToArray(),
                nr_genes = e1_n.ToArray(),
                genes = e1_s.ToArray(),
                best_genes_percentage = e1_p.ToArray()
            };

            element_rank e2 = new element_rank()
            {
                catName = catLabel2,
                p_fdr = e2_pv.ToArray(),
                average_fc = e2_fc.ToArray(),
                mad_fc = e2_m.ToArray(),
                nr_genes = e2_n.ToArray(),
                genes = e2_s.ToArray(),
                best_genes_percentage = e2_p.ToArray()
            };

            element_rank e3 = new element_rank()
            {
                catName = catLabel3,
                p_fdr = e3_pv.ToArray(),
                average_fc = e3_fc.ToArray(),
                mad_fc = e3_m.ToArray(),
                nr_genes = e3_n.ToArray(),
                genes = e3_s.ToArray(),
                best_genes_percentage = e3_p.ToArray()
            };

            element_rank e4 = new element_rank()
            {
                catName = catLabel4,
                p_fdr = e4_pv.ToArray(),
                average_fc = e4_fc.ToArray(),
                mad_fc = e4_m.ToArray(),
                nr_genes = e4_n.ToArray(),
                genes = e4_s.ToArray(),
                best_genes_percentage = e4_p.ToArray()
            };


            element_rank e5 = new element_rank()
            {
                catName = catLabel5,
                p_fdr = e5_pv.ToArray(),
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

        private List<element_rank> CreateVolcanoPlotData(element_fc theElements, int maxExtreme=-1)
        {

            //List<summaryInfo> all_elements = SortedElements(theElements.All, mode: SORTMODE.CATNAME, descending: false);
            //List<summaryInfo> pos_elements = SortedElements(theElements.Activated, mode: SORTMODE.CATNAME, descending: false);
            //List<summaryInfo> neg_elements = SortedElements(theElements.Repressed, mode: SORTMODE.CATNAME, descending: false);
            List<summaryInfo> best_elements = null;
            (gBestTable, best_elements) = BestElementScore(theElements);

            return BubblePlotData(best_elements,volcanoPlot:true, maxExtreme:maxExtreme);


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
