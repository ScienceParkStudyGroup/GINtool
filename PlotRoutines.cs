using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GINtool
{
    public static class PlotRoutines
    {
        public static Excel.Application theApp = null;

        static double estimatedFontSize(int nritems)
        {
            return -Math.Log10(nritems) + 2.6021;
        }

        static int fontsize(int nritems)
        {
            int size = (int)Math.Pow(10, estimatedFontSize(nritems));
            if (size < 2)
                size = 2;
            if (size > 10)
                size = 10;
            return size;
        }

        public static Excel.Chart CreateDistributionPlot(List<double> sortedFC, List<int> sortedIndex, string chartName)
        {
            if (theApp == null)
                return null;

            Excel.Worksheet aSheet = theApp.Worksheets.Add();

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
            Excel.Chart chartPage = myChart.Chart;

            //chartPage.ChartType = Excel.XlChartType.xlXYScatter;
            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

            var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

            var aSerie = series.NewSeries();
            aSerie.Name = String.Format("Serie {0}", 1);
            //aSerie.ChartType = Excel.XlChartType.xlXYScatter;


            //aSerie.XValues = sortedIndex.ToArray();
            aSerie.Values = sortedFC.ToArray();


            //distributionChart.Palette = (ChartColorPalette)Properties.Settings.Default.defaultPalette;
            //distributionChart.Titles.Add("distribution plot");

            var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            yAxis.HasTitle = true;
            yAxis.AxisTitle.Text = "index (sorted)";

            var xAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            xAxis.HasTitle = true;
            xAxis.AxisTitle.Text = "fold-change";


            chartPage.Axes(Excel.XlAxisType.xlCategory).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
            chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();
            chartPage.Legend.Delete();
            chartPage.ChartTitle.Delete();

            chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, chartName);


            aSheet.Delete();

            return chartPage;



        }


        public static Excel.Chart CreateRankingPlot(Excel.Worksheet aSheet, List<element_rank> element_Ranks, string chartName)
        {
            if (theApp == null)
                return null;


            string sheetName = chartName.Replace("Plot_", "Tab_");
            aSheet.Name = sheetName;


            Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
            Excel.Chart chartPage = myChart.Chart;

            chartPage.ChartType = Excel.XlChartType.xlXYScatter;

            var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

            int offset = 3;
            List<int> __offset = new List<int>() { offset };

            int[] _offset = element_Ranks.Select(w => offset += w.nr_genes.Length).ToArray();
            __offset.AddRange(_offset);
            int[] offsets = __offset.GetRange(0, element_Ranks.Count).ToArray();



            for (int i = element_Ranks.Count - 1; i >= 0; i--)
            {
                element_rank eRank = element_Ranks[i];
                var xy1 = series.NewSeries();
                xy1.Name = eRank.catName;
                xy1.ChartType = Excel.XlChartType.xlBubble3DEffect;

                int nrGenes = eRank.nr_genes.Length;
                if (eRank.mad_fc != null && nrGenes > 0)
                {
                    xy1.XValues = string.Format("={0}!$C${1}:$C${2}", sheetName, offsets[i], nrGenes + offsets[i] - 1); //element_rank.value.average_fc;
                    xy1.Values = string.Format("={0}!$D${1}:$D${2}", sheetName, offsets[i], nrGenes + offsets[i] - 1);  //element_rank.value.mad_fc;
                    xy1.BubbleSizes = string.Format("={0}!$B${1}:$B${2}", sheetName, offsets[i], nrGenes + offsets[i] - 1); //element_rank.value.nr_genes;

                    xy1.HasDataLabels = true;
                    dynamic dataLabels = xy1.DataLabels();
                    dataLabels.Format.TextFrame2.TextRange.InsertChartField(MsoChartFieldType.msoChartFieldRange, string.Format("={0}!$A${1}:$A${2}", sheetName, offsets[i], nrGenes + offsets[i] - 1));

                    dataLabels.ShowRange = true;
                    dataLabels.ShowValue = false;

                    offset += nrGenes;
                }
            }

            //foreach (var element_rank in element_Ranks.Select((value, index) => new { value, index }))
            //{

            //    }
            //}             

            chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
            chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();

            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
            chartPage.Legend.Delete();

            chartPage.ChartColor = 22;
            chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, chartName);

            return chartPage;

        }



        public static Excel.Chart CreateRankingPlot2(List<element_rank> element_Ranks, string chartName, bool bestPlot = false, bool bestNew = false)
        {
            if (theApp == null)
                return null;


            if (theApp == null)
                return null;

            Excel.Worksheet aSheet = theApp.Worksheets.Add();


            //string sheetName = chartName.Replace("PlotBest_", "TabBest_");
            //aSheet.Name = sheetName;


            Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
            Excel.Chart chartPage = myChart.Chart;

            chartPage.ChartType = Excel.XlChartType.xlXYScatter;

            var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

            int offset = 3;
            List<int> __offset = new List<int>() { offset };

            int[] _offset = element_Ranks.Select(w => offset += w.nr_genes.Length).ToArray();
            __offset.AddRange(_offset);
            int[] offsets = __offset.GetRange(0, element_Ranks.Count).ToArray();



            for (int i = element_Ranks.Count - 1; i >= 0; i--)
            {
                element_rank eRank = element_Ranks[i];
                var xy1 = series.NewSeries();
                xy1.Name = eRank.catName;
                xy1.ChartType = Excel.XlChartType.xlBubble3DEffect;

                int nrGenes = eRank.nr_genes.Length;
                if (eRank.mad_fc != null && nrGenes > 0)
                {
                    if (bestPlot & bestNew)
                    {
                        xy1.XValues = eRank.best_genes_percentage;
                        xy1.Values = eRank.average_fc;
                    }
                    else
                    {
                        xy1.XValues = eRank.average_fc;
                        xy1.Values = eRank.mad_fc;
                    }
                    xy1.BubbleSizes = eRank.nr_genes;

                    xy1.HasDataLabels = true;
                    dynamic dataLabels = xy1.DataLabels();

                    for (int g = 0; g < eRank.nr_genes.Length; g++)
                    {
                        Excel.Point _point = xy1.Points(g + 1);
                        _point.DataLabel.Text = eRank.genes[g];
                    }

                    dataLabels.ShowRange = true;
                    dataLabels.ShowValue = false;

                    offset += nrGenes;
                }
            }

            chartPage.Axes(Excel.XlAxisType.xlValue).HasTitle = true;
            chartPage.Axes(Excel.XlAxisType.xlCategory).HasTitle = true;

            string xLabel = (bestPlot & bestNew) ? "% logical regulation" : (bestPlot ? "average signed FC" : "average FC");
            string yLabel = (bestPlot & bestNew) ? "average FC" : (bestPlot ? "mad abs(FC)" : "mad FC");


            chartPage.Axes(Excel.XlAxisType.xlCategory).AxisTitle.Text = xLabel;
            chartPage.Axes(Excel.XlAxisType.xlValue).AxisTitle.Text = yLabel;


            chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
            chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();

            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
            chartPage.Legend.Delete();

            chartPage.ChartColor = (bestPlot | bestNew) ? 21 : 22;
            chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, chartName);

            aSheet.Delete();

            return chartPage;

        }



        public static Excel.Chart CreateCategoryPlot(element_fc element_Fcs, string chartName)
        {
            if (theApp == null)
                return null;

            Excel.Worksheet aSheet = theApp.Worksheets.Add();

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
            Excel.Chart chartPage = myChart.Chart;

            chartPage.ChartType = Excel.XlChartType.xlXYScatter;

            var series = (Excel.SeriesCollection)chartPage.SeriesCollection();

            int nrCategories = element_Fcs.All.Count;

            double MMAX = 0;
            double MMIN = 0;

            for (int _i = 0; _i < nrCategories; _i++)
            {
                if (element_Fcs.All[_i].p_values != null && element_Fcs.All[_i].p_values.Length > 0)
                {
                    if (element_Fcs.All[_i].fc_values.Min() < MMIN)
                        MMIN = element_Fcs.All[_i].fc_values.Min();
                    if (element_Fcs.All[_i].fc_values.Max() > MMAX)
                        MMAX = element_Fcs.All[_i].fc_values.Max();
                }
            }

            foreach (var element_Fc in element_Fcs.All.Select((value, index) => new { value, index }))
            {
                var xy1 = series.NewSeries();
                xy1.Name = element_Fc.value.catName;
                xy1.ChartType = Excel.XlChartType.xlXYScatter;
                if (element_Fc.value.fc_values != null && element_Fc.value.fc_values.Length > 0)
                {
                    xy1.XValues = element_Fc.value.fc_values;
                    xy1.Values = Enumerable.Repeat(element_Fc.index + 0.5, element_Fc.value.fc_values.Length).ToArray();
                    xy1.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
                    xy1.MarkerSize = 2;
                    //xy1.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, 0.1);
                    xy1.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, 0.4);
                    Excel.ErrorBars errorBars = xy1.ErrorBars;
                    errorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap;
                    errorBars.Format.Line.Weight = 1.25f;

                    // give each serie different color
                    switch (element_Fc.index % 6)
                    {
                        case 0:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent1;
                            break;
                        case 1:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent2;
                            break;
                        case 2:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent3;
                            break;
                        case 3:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent4;
                            break;
                        case 4:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent5;
                            break;
                        case 5:
                            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent6;
                            break;
                    }


                    //Excel.Points sPoints = xy1.Points();

                    ////xy1.HasDataLabels = true;

                    ////For every row in the values table, plot the date against the variable value
                    //for (int p = 0; p < element_Fc.value.fc.Count(); p++)
                    //{
                    //    Excel.Point lPoint = sPoints.Item(p + 1);
                    //    //lPoint.Name = "P" + p.ToString();

                    //    //myChart.Series[Variable].Points.AddXY(Convert.ToDateTime(row["Date"].ToString()), row["Variable"].ToString());
                    //    lPoint.HasDataLabel = true;
                    //    lPoint.DataLabel.Text = "P" + p.ToString(); // " = #VALY \r\nDate = #VALX{d} \r\np = "+ p.ToString();
                    //    lPoint.DataLabel.Font.Size = 2;
                    //    //Excel.Point lPoint = sPoints.Item(p);
                    //    //points += 1;
                    //}

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
                    xy2.DataLabels(_i + 1).Text = element_Fcs.All[_i].catName;
                }

                xy2.DataLabels().Position = Excel.XlDataLabelPosition.xlLabelPositionLeft;
                xy2.DataLabels().Font().Size = fontsize(element_Fcs.All.Count);

            }


            chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
            chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();

            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
            chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
            chartPage.Axes(Excel.XlAxisType.xlValue).MaximumScale = nrCategories;
            chartPage.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
            chartPage.Legend.Delete();

            chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, chartName);

            aSheet.Delete();

            return chartPage;

        }



        public static (float, float) CalculateStepSize(float range, float targetSteps)
        {
            // calculate an initial guess at step size
            float tempStep = range / targetSteps;

            // get the magnitude of the step size
            float mag = (float)Math.Floor(Math.Log10(tempStep));
            float magPow = (float)Math.Pow(10, mag);

            // calculate most significant digit of the new step size
            float magMsd = (int)(tempStep / magPow + 0.5);

            // promote the MSD to either 1, 2, or 5
            if (magMsd > 5.0)
                magMsd = 10.0f;
            else if (magMsd > 2.0)
                magMsd = 5.0f;
            else if (magMsd > 1.0)
                magMsd = 2.0f;

            return (magMsd * magPow, magMsd);
        }
    }


}