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
    public partial class GinRibbon
    {
       
        /// <summary>
        /// Create the distribution plot
        /// </summary>
        /// <param name="aOutput"></param>        
        // private void DistributionPlot(List<FC_BSU> aOutput)
        private void DistributionPlot(List<BsuLinkedItems> aOutput)
        {
            gApplication.EnableEvents = false;
            gApplication.DisplayAlerts = false;

            SysData.DataTable _fc_BSU_ = CreateRegulonUsageTable(aOutput);
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
        /// <param name="cat_Elements">The selected elements from the dialog</param>
        /// <param name="topTenFC"></param>
        /// <param name="topTenP"></param>
        /// <param name="outputTable"></param>
        //private void SpreadingPlot(List<FC_BSU> aOutput, SysData.DataTable aSummary, List<cat_elements> cat_Elements, int topTenFC = -1, int topTenP = -1, bool outputTable = false)
        private void SpreadingPlot(List<cat_elements> cat_Elements, int topTenFC = -1, int topTenP = -1, bool outputTable = false)
        {
            
            AddTask(TASKS.CATEGORY_CHART);

            Func<DataView, List<cat_elements>, int, int, element_fc> CatElementsPtr = null;        

            ////SysData.DataTable _fc_BSU = ReformatRegulonResults(aOutput);
            //if (gRegulonTable is null | NeedsUpdate(UPDATE_FLAGS.TRegulon))
            //{
            //    gRegulonTable = CreateRegulonUsageTable(GetDataSelection());
                

            //if (gCategoryTable is null | NeedsUpdate(UPDATE_FLAGS.TCategory))
            //    gRegulonTable = CreateRegulonUsageTable(GetDataSelection());

            //// SysData.DataTable _fc_BSU = ReformatRegulonResults(aOutput);
            cat_Elements = GetUniqueElements(cat_Elements);

            // HashSet ensures unique list
            HashSet<string> lRegulons = new HashSet<string>();

            foreach (SysData.DataRow row in gRegulonTable.Rows)
                lRegulons.Add(row.ItemArray[0].ToString());

            SysData.DataView dataView = gRegulonTable.AsDataView();
            element_fc catPlotData;
            //if (Properties.Settings.Default.useCat)
            //{
            //    //funcPtr = CatElements2ElementsFC(dataView, cat_Elements, topTenFC, topTenP);
            // CatElementsPtr = Properties.Settings.Default.useCat ? CatElements2ElementsFC : Regulons2ElementsFC;
            //}
            //else
            //    //catPlotData = Regulons2ElementsFC(dataView, cat_Elements, topTenFC: topTenFC, topTenP: topTenP); // need to alter caller
            //    CatElementsPtr = Regulons2ElementsFC;// need to alter caller

            if (Properties.Settings.Default.useCat)
                CatElementsPtr = CatElements2ElementsFC;
            else
                CatElementsPtr = Regulons2ElementsFC;
            
            catPlotData = CatElementsPtr(dataView, cat_Elements, topTenFC, topTenP);

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
        private void RankingPlot(List<cat_elements> cat_Elements, SysData.DataTable aSummary)
        {
            AddTask(TASKS.REGULON_CHART);

            //SysData.DataTable _fc_BSU = ReformatRegulonResults(aOutput);

            cat_Elements = GetUniqueElements(cat_Elements);

            // HashSet ensures unique list
            //HashSet<string> lRegulons = new HashSet<string>();

            //foreach (SysData.DataRow row in aSummary.Rows)
            //    lRegulons.Add(row.ItemArray[0].ToString());

            SysData.DataView dataView = aSummary.AsDataView();
            element_fc catPlotData;
            if (Properties.Settings.Default.useCat)
            {
                catPlotData = CatElements2ElementsFC(dataView, cat_Elements);
            }
            else
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements);

            (List<element_rank> plotData, List<summaryInfo> _all, List<summaryInfo> _pos, List<summaryInfo> _neg, List<summaryInfo> _best) = CreateRankingPlotData(catPlotData);

            int suffix = 0;
            
            if (gSettings.useCat)
                suffix = FindSheetNames(new string[] { "CatRankPlot_", "Plot_", "PlotBest_", "CatRankTable_" });
            else
                suffix = FindSheetNames(new string[] { "RegRankPlot_", "Plot_", "PlotBest_", "RegRankTable_" });


            //int chartNr = Properties.Settings.Default.useCat ? NextWorksheet("CatRankPlot_") : NextWorksheet("RegRankPlot_");
            string chartName = (Properties.Settings.Default.useCat ? "CatRankPlot_" : "RegRankPlot_") + suffix.ToString();
            string chartNameBest = chartName.Replace("Plot_", "PlotBest_");
            
            
            CreateRankingDataSheet(catPlotData, _all, _pos, _neg, _best,suffix);

            PlotRoutines.CreateRankingPlot2(plotData, chartName);

            if (!(_best is null))
            {
                List<element_rank> _bestRankData = BubblePlotData(_best);
                PlotRoutines.CreateRankingPlot2(_bestRankData, chartNameBest + "_v1", bestPlot: true);
                PlotRoutines.CreateRankingPlot2(_bestRankData, chartNameBest + "_v2", bestPlot: true, bestNew: true);
            }


            this.RibbonUI.ActivateTab("TabGINtool");

            RemoveTask(TASKS.REGULON_CHART);

        }
    }
}
