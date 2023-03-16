﻿using System;
using System.Collections.Generic;
using System.Data;
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

            //SysData.DataTable _fc_BSU_ = CreateRegulonUsageTable(aOutput);
            SysData.DataTable _fc_BSU_ = CreateGeneUsageTable(aOutput);



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
        private void SpreadingPlot(List<cat_elements> cat_Elements, int topTenFC = -1, bool outputTable = false)
        {

            AddTask(TASKS.CATEGORY_CHART);

            Func<DataView, List<cat_elements>, int, element_fc> CatElementsPtr = null;
            
            cat_Elements = GetUniqueElements(cat_Elements);

            SysData.DataView dataView = gSettings.useCat ? gCategoryTable.AsDataView() : gRegulonTable.AsDataView();
            element_fc catPlotData;

            if (Properties.Settings.Default.useCat)
                CatElementsPtr = CatElements2ElementsFC;
            else
                CatElementsPtr = Regulons2ElementsFC;

            catPlotData = CatElementsPtr(dataView, cat_Elements, topTenFC);

            string postFix = topTenFC > -1 ? string.Format("Top{0}FC", topTenFC) : "";
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

        
        private void VolcanoPlot(List<cat_elements> cat_Elements, SysData.DataTable aSummary, int maxExtreme=-1)
        {
            AddTask(TASKS.VOLCANO_PLOT);

            cat_Elements = GetUniqueElements(cat_Elements);

            SysData.DataView dataView = aSummary.AsDataView();
            element_fc catPlotData;
            if (Properties.Settings.Default.useCat)
            {
                catPlotData = CatElements2ElementsFC(dataView, cat_Elements);
            }
            else
                catPlotData = Regulons2ElementsFC(dataView, cat_Elements);

            List<element_rank> plotData = CreateVolcanoPlotData(catPlotData, maxExtreme:maxExtreme);
            int suffix = 0;

            if (gSettings.useCat)
                suffix = FindSheetNames(new string[] { "CatVolcanoPlot", "Plot"});
            else
                suffix = FindSheetNames(new string[] { "RegVolcanoPlot", "Plot"});

            string chartName = (Properties.Settings.Default.useCat ? "CatVolcanoPlot_" : "RegVolcanoPlot_") + suffix.ToString();            
            PlotRoutines.CreateVolcanoPlot(plotData, chartName);

            this.RibbonUI.ActivateTab("TabGINtool");
            RemoveTask(TASKS.VOLCANO_PLOT);

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
                suffix = FindSheetNames(new string[] { "CatRankPlot", "Plot", "CatRankPlotBest_v1", "CatRankPlotBest_v2", "CatRankTable" });
            else
                suffix = FindSheetNames(new string[] { "RegRankPlot", "Plot", "RegRankPlotBest_v1", "RegRankPlotBest_v2", "RegRankTable" });


            //int chartNr = Properties.Settings.Default.useCat ? NextWorksheet("CatRankPlot_") : NextWorksheet("RegRankPlot_");
            string chartName = (Properties.Settings.Default.useCat ? "CatRankPlot_" : "RegRankPlot_") + suffix.ToString();
            string chartNameBestv1 = chartName.Replace("Plot_", "PlotBest_v1_");
            string chartNameBestv2 = chartName.Replace("Plot_", "PlotBest_v2_");


            CreateRankingDataSheet(catPlotData, _all, _pos, _neg, _best, suffix);

            PlotRoutines.CreateRankingPlot2(plotData, chartName);

            if (!(_best is null))
            {
                List<element_rank> _bestRankData = BubblePlotData(_best);
                PlotRoutines.CreateRankingPlot2(_bestRankData, chartNameBestv1, bestPlot: true);
                PlotRoutines.CreateRankingPlot2(_bestRankData, chartNameBestv2, bestPlot: true, bestNew: true);
            }


            this.RibbonUI.ActivateTab("TabGINtool");

            RemoveTask(TASKS.REGULON_CHART);

        }
    }
}
