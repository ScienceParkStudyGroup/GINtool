using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;
using System.IO;
using SysData = System.Data;
using System.Data;
using GINtool.Properties;
using Microsoft.Office.Core;
using System.Globalization;
using System.Threading;

namespace GINtool
{
    public static class PlotRoutines
    {
        //Chart distributionChart = null;
        //Chart enrichmentChart = null;
        //Chart scoreChart = null;
        //Chart qPlot = null;
        //ChartColorPalette thePallete = ChartColorPalette.Excel;
        public static Excel.Application theApp = null;
        //List<float> sortedGenesValues;
        //List<int> sortedGenesInt;

        //int multiplier = 25;
        //int nrRegulons = 3;
        //int nrGenes = 250;

        //public void UpdateFigures()
        //{
        //    SetNrRegulons(nrRegulons);
        
        //}

        //public PlotRoutines(Excel.Application app)
        //{
        //    //thePallete = (ChartColorPalette)Properties.Settings.Default.defaultPalette;
        //    theApp = app;
        //}

        ////public void SetPalette(ChartColorPalette pal)
        ////{
        ////    thePallete = pal;
        ////}

        //public void SetNrRegulons(int nr)
        //{
        //    nrRegulons = nr;
        //    //int multiplier = 25;
        //    //if (nr <= 10) multiplier = 25;
        //    if (nr > 10 & nr <= 20) multiplier = 20;
        //    if (nr > 20) multiplier = 15;
            

        //}

        //public void SetNrGenes(int nr)
        //{
        //    nrGenes = nr;
        //}

        //private int[] CumulativeSums(int[] values)
        //{
        //    if (values == null || values.Length == 0) return new int[0];

        //    var results = new int[values.Length];
        //    results[0] = values[0];

        //    for (var i = 1; i < values.Length; i++)
        //    {
        //        results[i] = results[i - 1] + values[i];
        //    }

        //    return results;
        //}

        //public Excel.Shape CreateEmfShape(Chart chart, IntPtr hWnd, Excel.Application theApp)
        //{
        //    MemoryStream theFile = new MemoryStream();
        //    chart.SaveImage(theFile, ChartImageFormat.EmfPlus);
        //    theFile.Flush();
        //    theFile.Position = 0;
        //    Metafile OriginalImage = (Metafile)Metafile.FromStream(theFile, true);
        //    EmfHelper.PutEnhMetafileOnClipboard(hWnd, OriginalImage);
        //    theFile.Close();
        //    theFile.Dispose();
        //    Excel.Worksheet _tmpSheet = theApp.ActiveWorkbook.Sheets.Add();
        //    _tmpSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
        //    Excel.Range aRange = _tmpSheet.Cells[4, 4];
        //    _tmpSheet.Paste(aRange);

        //    List<Excel.Shape> lShapes = new List<Excel.Shape>();

        //    foreach (Excel.Shape aShape in _tmpSheet.Shapes)
        //    {
        //        lShapes.Add(aShape);
        //    }

        //    return lShapes[0];

        //}


        //public Excel.Shape DrawDistributionChart()
        //{
        //    if (theApp == null)
        //        return null;
            
        //    IntPtr hWnd = (IntPtr)theApp.Hwnd;

        //    return null;

        //}

        //public Excel.Shape DrawDistributionPlot(HashSet<string> aRegulons, SysData.DataTable aTable, Excel.Worksheet aSheet)
        //{
        //    if (theApp == null)
        //        return null;

        //    IntPtr hWnd = (IntPtr)theApp.Hwnd;

        //    nrGenes = aTable.Rows.Count;
        //    nrRegulons = aRegulons.Count;

        //    distributionChart = new Chart();
        //    distributionChart.Legends.Clear();
        //    ChartArea chartArea = new ChartArea("distibutionChart");

        //    distributionChart.Height = 300; // nrRegulons * multiplier;
        //    distributionChart.Width = 700;

        //    chartArea.AxisX.MajorGrid.Enabled = false;
        //    chartArea.AxisY.MajorGrid.Enabled = false;

        //    chartArea.AxisX.Minimum = 0;
        //    chartArea.AxisX.Maximum = nrGenes;
            

        //    distributionChart.ChartAreas.Add(chartArea);
          

        //    List<float> _values = new List<float>();
        //    foreach (SysData.DataRow row in aTable.Rows)
        //    {
        //        _values.Add(row.Field<float>("FC"));
        //    }

        //    float[] __values = _values.ToArray();

        //    var sortedGenes = __values
        //        .Select((x, i) => new KeyValuePair<float, int>(x, i))
        //        .OrderBy(x => x.Key)
        //        .ToList();

        //    sortedGenesValues = sortedGenes.Select(x => x.Key).ToList();
        //    sortedGenesInt = sortedGenes.Select(x => x.Value).ToList();

        //    string toClipboard = string.Join("\n", sortedGenesInt.ToArray());
        //    System.Windows.Forms.Clipboard.SetData(System.Windows.Forms.DataFormats.Text,toClipboard);

        //    System.Windows.Forms.DataVisualization.Charting.Series aSerie = distributionChart.Series.Add(String.Format("Serie {0}", 1));
        //    aSerie.ChartType = SeriesChartType.Column;

        //    for (int _p = 0; _p < sortedGenesValues.Count; _p++)
        //    {
        //        aSerie.Points.AddXY(_p, sortedGenesValues[_p]);
        //    }

        //    distributionChart.Palette = (ChartColorPalette)Properties.Settings.Default.defaultPalette; 
        //    distributionChart.Titles.Add("distribution plot");

        //    chartArea.AxisX.Title = "index (sorted)";
        //    chartArea.AxisY.Title = "fold change";

            

        //    return CreateEmfShape(distributionChart, hWnd, theApp);

        //}

        //public Excel.Shape DrawEnrichmentScoreChart(HashSet<string> aRegulons, SysData.DataTable aTable)
        //{
        //    return null;
        //}


        //private void SetAxisTickLabels(Axis aAxis, double amin, double amax)
        //{
        //    aAxis.CustomLabels.Clear();

        //    int nrsteps = 5;
        //    (float stepsize, float mag) = CalculateStepSize((float)(amax - amin), nrsteps);

        //    double lMin = Math.Floor(amin);

        //    int lMag = (int)Math.Round(Math.Log10(stepsize));
        //    if (lMag < 0)
        //        lMag = Math.Abs(lMag);
        //    else
        //        lMag = 0;


        //    for (int i = 0; i < nrsteps; i++)
        //    {
        //        aAxis.CustomLabels.Add(lMin + (i * stepsize), lMin + ((i + 1) * stepsize), String.Format("{0:F"+lMag+"}", lMin + (i * stepsize)), 0, LabelMarkStyle.SideMark, GridTickTypes.TickMark);
        //    }

        //    aAxis.Maximum = lMin + (nrsteps * stepsize);
        //    aAxis.Minimum = lMin;

        //}

        //public Excel.Shape DrawQPlot(HashSet<string> aRegulons, SysData.DataTable aTable)
        //{
        //    if (theApp == null)
        //        return null;

        //    IntPtr hWnd = (IntPtr)theApp.Hwnd;

        //    nrGenes = aTable.Rows.Count;
        //    nrRegulons = aRegulons.Count;



        //    Chart qPlot = new Chart();
        //    qPlot.Legends.Clear();
        //    ChartArea chartArea = new ChartArea("qPlot");

        //    qPlot.Height = 800;
        //    qPlot.Width = 900;

        //    SysData.DataView dataView = aTable.AsDataView();
        //    double MMAX_X = (double)(float)aTable.Rows[0]["Pvalue"];
        //    double MMIN_X = (double)(float)aTable.Rows[0]["Pvalue"];

        //    double MMAX_Y = (double)(float)aTable.Rows[0]["FC"];
        //    double MMIN_Y = (double)(float)aTable.Rows[0]["FC"];


        //    System.Windows.Forms.DataVisualization.Charting.Legend legend = new Legend("Series");
        //    qPlot.Legends.Add(legend);



        //    foreach (string regulon in aRegulons)
        //    {
        //        System.Windows.Forms.DataVisualization.Charting.Series serie = qPlot.Series.Add(String.Format("{0}", regulon));
        //        serie.ChartType = SeriesChartType.Point;

        //        serie.LegendText = regulon;
        //        serie.IsVisibleInLegend = true;

        //        dataView.RowFilter = String.Format("Regulon = '{0}'", regulon);
        //        List<float> fc = new List<float>();

        //        SysData.DataTable dataTable = dataView.ToTable();
        //        int nrRows = dataTable.Rows.Count;
        //        float[] vs = new float[nrRows];
        //        for (int _r = 0; _r < nrRows; _r++)
        //        {
        //            double _y = (double)(float)dataTable.Rows[_r]["FC"];
        //            double _x = (double)Math.Log10((float)dataTable.Rows[_r]["Pvalue"]);
        //            if (_y > MMAX_Y) { MMAX_Y = _y; }
        //            if (_y < MMIN_Y) { MMIN_Y = _y; }
        //            if (_x > MMAX_X) { MMAX_X = _x; }
        //            if (_x < MMIN_X) { MMIN_X = _x; }

        //            serie.Points.AddXY(_x, _y);

        //        }                                

        //    }

        //    qPlot.ChartAreas.Add(chartArea);
        //    qPlot.Titles.Add("fold change vs. p-values");
        //    chartArea.AxisX.Title = "fold change";
        //    chartArea.AxisY.Title = "log10(p-value)";

        //    SetAxisTickLabels(chartArea.AxisX, MMIN_X, MMAX_X);
        //    SetAxisTickLabels(chartArea.AxisY, MMIN_Y, MMAX_Y);


        //    chartArea.AxisX.MajorGrid.Enabled = false;
        //    chartArea.AxisY.MajorGrid.Enabled = false;


        //     qPlot.Palette = (System.Windows.Forms.DataVisualization.Charting.ChartColorPalette)Properties.Settings.Default.defaultPalette;

        //    return CreateEmfShape(qPlot, hWnd, theApp);

        //}

        public static Excel.Chart CreateCategoryPlot(List<element_fc> element_Fcs)
        {
            if (theApp == null)
                return null;

            Excel.Worksheet aSheet = theApp.Worksheets.Add();

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

            foreach(var element_Fc in element_Fcs.Select((value,index) => new {value, index}))
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



        //public (Excel.Worksheet,Excel.Chart) CreateExcelChart(List<element_fc> element_Fcs)
        //{
        //    if (theApp == null)
        //        return (null, null);

           
            
        //    Excel.Worksheet aSheet = theApp.Worksheets.Add();

        //    var missing = System.Type.Missing;

        //    //Excel.ChartObject myChart = (Excel.ChartObject)theApp.Charts.Add2(Type.Missing, Type.Missing, Type.Missing,Type.Missing);

        //    Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
        //    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
        //    Excel.Chart chartPage = myChart.Chart;

        //    chartPage.ChartType = Excel.XlChartType.xlXYScatter;


        //    return (null, null);




        //    //ChartColorPalette pal = (ChartColorPalette)Properties.Settings.Default.defaultPalette;









        //    //nrGenes = aTable.Rows.Count;
        //    //nrRegulons = aRegulons.Count;

        //    //List<float[]> fc = new List<float[]>();

        //    //SysData.DataView dataView = aTable.AsDataView();

        //    //double MMAX = (double)(float)aTable.Rows[0]["FC"];
        //    //double MMIN = (double)(float)aTable.Rows[0]["FC"];


        //    //foreach (string regulon in aRegulons)
        //    //{
        //    //    dataView.RowFilter = String.Format("Regulon = '{0}'", regulon);
        //    //    SysData.DataTable dataTable = dataView.ToTable();
        //    //    int nrRows = dataTable.Rows.Count;
        //    //    float[] vs = new float[nrRows];
        //    //    int[] ys = new int[nrRows];
        //    //    for (int _r = 0; _r < nrRows; _r++)
        //    //    {
        //    //        double _val = (double)(float)dataTable.Rows[_r]["FC"];
        //    //        if (_val > MMAX) { MMAX = _val; }
        //    //        if (_val < MMIN) { MMIN = _val; }
        //    //        vs[_r] = (float)_val;
        //    //        ys[_r] = fc.Count;

        //    //    }

        //    //    fc.Add(vs);
        //    //}




        //    //var series = (Excel.SeriesCollection)chartPage.SeriesCollection();
        //    //for (int i = 0; i < nrRegulons; i++)
        //    //{
        //    //    var xy1 = series.NewSeries();
        //    //    xy1.Name = aRegulons.ToArray()[i]; // aRstring.Format("Regulon {0}", i + 1);
        //    //    xy1.ChartType = Excel.XlChartType.xlXYScatter;
        //    //    xy1.XValues = fc[i];

        //    //    xy1.Values = Enumerable.Repeat(i + 0.5, fc[i].Length).ToArray();

        //    //    xy1.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
        //    //    xy1.MarkerSize = 2;
        //    //    xy1.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, 0.1);
        //    //    Excel.ErrorBars errorBars = xy1.ErrorBars;
        //    //    errorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap;
        //    //    errorBars.Format.Line.Weight = 1.25f;


        //    //    switch (i % 6)
        //    //    {
        //    //        case 0:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent1;
        //    //            break;
        //    //        case 1:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent2;
        //    //            break;
        //    //        case 2:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent3;
        //    //            break;
        //    //        case 3:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent4;
        //    //            break;
        //    //        case 4:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent5;
        //    //            break;
        //    //        case 5:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent6;
        //    //            break;
        //    //    }

        //    //    var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
        //    //    //yAxis.AxisTitle.Text = "Regulon";
        //    //    Excel.TickLabels labels = yAxis.TickLabels;
        //    //    labels.Offset = 1;

        //    //}


        //    //chartPage.ChartColor = 1; // Properties.Settings.Default.defaultPalette;

        //    //// as a last step, add the axis labels series

        //    //if (true)
        //    //{ 

        //    //    var xy2 = series.NewSeries();

        //    //    xy2.ChartType = Excel.XlChartType.xlXYScatter;
        //    //    //# Excel.Range rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i*2)+1], _tmpSheet.Cells[6, (i * 2) + 1]];
        //    //    xy2.XValues = Enumerable.Repeat(MMIN, nrRegulons).ToArray();

        //    //    //rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i * 2) + 2], _tmpSheet.Cells[6, (i * 2) + 2]];
        //    //    float[] yv = new float[nrRegulons];
        //    //    for (int i = 0; i < nrRegulons; i++)
        //    //    {
        //    //        yv[i] = ((float)i) + 0.5f;
        //    //    }

        //    //    xy2.Values = yv;

        //    //    xy2.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

        //    //    //string[] squares = Enumerable.Range(1, nrRegulons).Select(x => string.Format("Regulon {0}", x)).ToArray();

        //    //    xy2.HasDataLabels = true;

        //    //    //xy2.DataLabels(0).Text = "test";
        //    //    //l0.Text = "pietje";
        //    //    //int j = 0;

        //    //    for (int i = 0; i < nrRegulons; i++)
        //    //    //foreach (string regulon in aRegulons)
        //    //    {

        //    //            xy2.DataLabels(i + 1).Text = aRegulons.ToArray()[i] ;
        //    //        //j++;

        //    //    }

        //    //    xy2.DataLabels().Position = Excel.XlDataLabelPosition.xlLabelPositionLeft;

        //    //}

        //    //chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();


        //    //chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).MaximumScale = aRegulons.Count;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
        //    //chartPage.Legend.Delete();

        //    //chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);
        //    ////myChart.Pal = (ChartColorPalette)Properties.Settings.Default.defaultPalette;

        //    ////aSheet.Delete();

        //    //return (aSheet,chartPage);
        //}



        //public void EnrichmentPlot(Excel.Worksheet aSheet, int sheetNr)
        //{
        //    if (theApp == null)
        //        return;


        //    theApp.EnableEvents = false;
        //    theApp.DisplayAlerts = false;
        //    theApp.ScreenUpdating =false;
            

        //    Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
        //    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);            
        //    Excel.Chart chartPage = myChart.Chart;


        //    Excel.Range wholeRange = (Excel.Range)aSheet.UsedRange;


        //    //aSheet.get_Range("a1").EntireRow.EntireColumn;


        //    chartPage.ChartType = Excel.XlChartType.xlXYScatter;
        //    chartPage.SetSourceData(wholeRange);

        //    //aSheet.Range[aSheet.Cells[1, 1], aSheet.Cells[totrow, nrRegulons + 1]]);


        //    foreach(Excel.Series serie in chartPage.SeriesCollection())
        //    {
        //        serie.MarkerSize = 2;
        //        serie.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
        //    }

            


        //    //CultureInfo MyCulture = new CultureInfo("en-US"); // your culture here 
        //    //Thread.CurrentThread.CurrentCulture = MyCulture;

        //    //Excel.Axis lAxis = chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
        //    //lAxis.HasTitle = true;
        //    //lAxis.AxisTitle.Text = "Fold Change";


        //    Excel.Axis tAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
        //    tAxis.HasTitle = true;
        //    tAxis.AxisTitle.Text = "Fold Change";

        //    chartPage.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

        //    chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, string.Format("CongruencePlot_{0}", sheetNr));


        //    //chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).AxisTitle.Text = "Fold Change";


        //    //var missing = System.Type.Missing;

        //    ////Excel.ChartObject myChart = (Excel.ChartObject)theApp.Charts.Add2(Type.Missing, Type.Missing, Type.Missing,Type.Missing);

        //    //Excel.ChartObjects xlCharts = (Excel.ChartObjects)aSheet.ChartObjects(Type.Missing);
        //    //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 500, 500);
        //    //Excel.Chart chartPage = myChart.Chart;

        //    //chartPage.ChartType = Excel.XlChartType.xlXYScatter;

        //    //var series = (Excel.SeriesCollection)chartPage.SeriesCollection();
        //    //for (int i = 0; i < nrRegulons; i++)
        //    //{
        //    //    var xy1 = series.NewSeries();
        //    //    xy1.Name = aRegulons.ToArray()[i]; // aRstring.Format("Regulon {0}", i + 1);
        //    //    xy1.ChartType = Excel.XlChartType.xlXYScatter;
        //    //    xy1.XValues = fc[i];

        //    //    xy1.Values = Enumerable.Repeat(i + 0.5, fc[i].Length).ToArray();

        //    //    xy1.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
        //    //    xy1.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypeFixedValue, 0.1);
        //    //    Excel.ErrorBars errorBars = xy1.ErrorBars;
        //    //    errorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap;
        //    //    errorBars.Format.Line.Weight = 1.25f;


        //    //    switch (i % 6)
        //    //    {
        //    //        case 0:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent1;
        //    //            break;
        //    //        case 1:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent2;
        //    //            break;
        //    //        case 2:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent3;
        //    //            break;
        //    //        case 3:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent4;
        //    //            break;
        //    //        case 4:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent5;
        //    //            break;
        //    //        case 5:
        //    //            errorBars.Format.Line.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent6;
        //    //            break;
        //    //    }

        //    //    var yAxis = (Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
        //    //    //yAxis.AxisTitle.Text = "Regulon";
        //    //    Excel.TickLabels labels = yAxis.TickLabels;
        //    //    labels.Offset = 1;

        //    //}


        //    //chartPage.ChartColor = 1; // Properties.Settings.Default.defaultPalette;

        //    //// as a last step, add the axis labels series

        //    //if (true)
        //    //{

        //    //    var xy2 = series.NewSeries();

        //    //    xy2.ChartType = Excel.XlChartType.xlXYScatter;
        //    //    //# Excel.Range rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i*2)+1], _tmpSheet.Cells[6, (i * 2) + 1]];
        //    //    xy2.XValues = Enumerable.Repeat(MMIN, nrRegulons).ToArray();

        //    //    //rng = (Excel.Range)_tmpSheet.Range[_tmpSheet.Cells[3, (i * 2) + 2], _tmpSheet.Cells[6, (i * 2) + 2]];
        //    //    float[] yv = new float[nrRegulons];
        //    //    for (int i = 0; i < nrRegulons; i++)
        //    //    {
        //    //        yv[i] = ((float)i) + 0.5f;
        //    //    }

        //    //    xy2.Values = yv;

        //    //    xy2.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

        //    //    //string[] squares = Enumerable.Range(1, nrRegulons).Select(x => string.Format("Regulon {0}", x)).ToArray();

        //    //    xy2.HasDataLabels = true;

        //    //    //xy2.DataLabels(0).Text = "test";
        //    //    //l0.Text = "pietje";
        //    //    //int j = 0;

        //    //    for (int i = 0; i < nrRegulons; i++)
        //    //    //foreach (string regulon in aRegulons)
        //    //    {

        //    //        xy2.DataLabels(i + 1).Text = aRegulons.ToArray()[i];
        //    //        //j++;

        //    //    }

        //    //    xy2.DataLabels().Position = Excel.XlDataLabelPosition.xlLabelPositionLeft;

        //    //}

        //    //chartPage.Axes(Excel.XlAxisType.xlValue).TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).MajorGridLines.Delete();


        //    //chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.Weight = 0.25;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).Format.Line.DashStyle = Excel.XlLineStyle.xlDashDot;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).MaximumScale = aRegulons.Count;
        //    //chartPage.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
        //    //chartPage.Legend.Delete();

        //    //chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);
        //    ////myChart.Pal = (ChartColorPalette)Properties.Settings.Default.defaultPalette;

        //    //aSheet.Delete();


        //    //Excel.Shape shp = aSheet.Shapes.AddChart(Excel.XlChartType.xlXYScatter, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //.Select(false);

        //    //Excel.Chart chart = shp.Chart;

        //    //chart.SetSourceData(aSheet.Range[aSheet.Cells[1, 1], aSheet.Cells[totrow, nrRegulons + 1]]);
        //    //shp.Name = "TotalPreseasonsSales";
        //    //chart.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);

        //    ///chartPage.Location(Excel.XlChartLocation.xlLocationAsNewSheet, Type.Missing);

        //    //chart.SetSourceData(aSheet.get_Range(string.Format("{0}2:{1}{2}", (char)('A' + (0 + 0)),

        //    //(char)('A' + 0 + 1), 0 + 2), Type.Missing), Type.Missing);



        //    theApp.EnableEvents = true;
        //    theApp.DisplayAlerts = true;
        //    theApp.ScreenUpdating = true;

        //    //return (aSheet, chartPage);
        //}


        //public void DrawEnrichmentChart(HashSet<string> aRegulons, SysData.DataTable aTable)
        //{
        //    if (theApp == null)
        //        return;


        //    Excel.Worksheet newSheet = theApp.Worksheets.Add();

        //    var missing = System.Type.Missing;

        //    Excel.Shape _shape = newSheet.Shapes.AddChart2(missing, Excel.XlChartType.xlXYScatter, missing, missing, missing, missing, missing);


        //    Excel._Chart ct = _shape.Chart;

        //    //Excel._Chart ct = (Excel._Chart)theApp.ThisWorkbook.Charts.Add(missing, missing, missing, missing);

        //    //Excel.Chart chart = theApp.Charts.Add(Excel.XlChartType.xlXYScatter);

        //    //Excel.Range oRng = newSheet.get_Range("A1");
        //    //Excel.Chart ct = newSheet.Shapes.AddChart().Chart;
           
            
        //    ct.ChartType = Excel.XlChartType.xlXYScatterSmooth;
        //    //ct.ChartWizard(oRng, Excel.XlChartType.xlXYScatterSmooth, missing, missing, missing, missing, missing, missing, "x axis", missing, missing);
        //    Excel.SeriesCollection theCollection = ct.SeriesCollection();
            
            
            
        //    nrGenes = aTable.Rows.Count;
        //    nrRegulons = aRegulons.Count;

        //    List<float[]> fc = new List<float[]>();

        //    SysData.DataView dataView = aTable.AsDataView();

        //    double MMAX = (double)(float)aTable.Rows[0]["FC"];
        //    double MMIN = (double)(float)aTable.Rows[0]["FC"];


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

        //        Excel.Series aSerie = theCollection.NewSeries();
        //        aSerie.XValues = vs;
        //        aSerie.Values = fc; 
        //        //string toClipboard = string.Join("\n", vs.ToArray());
        //        //System.Windows.Forms.Clipboard.SetData(System.Windows.Forms.DataFormats.Text, toClipboard);

        //        fc.Add(vs);
        //    }



        //    //enrichmentChart = new Chart();
        //    //enrichmentChart.Legends.Clear();
        //    //ChartArea chartArea = new ChartArea("enrichmentChart");

        //    //enrichmentChart.Height = nrRegulons * 40;
        //    //enrichmentChart.Width = 400;

        //    //chartArea.AxisX.MajorGrid.Enabled = false;
        //    //chartArea.AxisY.MajorGrid.Enabled = false;


        //    //chartArea.AxisY.Minimum = 0;
        //    //chartArea.AxisY.Maximum = nrRegulons + 1;

        //    ////enrichmentChart.ChartAreas.Add(chartArea);

        //    ////float stepsize = CalculateStepSize((float)(MMAX - MMIN), 5);

        //    ////double lMin = Math.Sign(MMIN) * Math.Ceiling(Math.Abs(MMIN));
        //    ////double lMax = Math.Sign(MMAX) * Math.Ceiling(Math.Abs(MMAX));


        //    ////for (int i = 0; i < 5; i++)
        //    ////{
        //    ////    chartArea.AxisX.CustomLabels.Add(lMin + (i * stepsize), lMin + ((i+1)* stepsize), String.Format("{0}", lMin + (i * stepsize)), 0, LabelMarkStyle.SideMark, GridTickTypes.TickMark);
        //    ////}

        //    ////chartArea.AxisX.Maximum = lMax;
        //    ////chartArea.AxisX.Minimum = lMin;


        //    //SetAxisTickLabels(chartArea.AxisX, MMIN, MMAX);


        //    //int nrRegulon = 0;

        //    //foreach (string regulon in aRegulons)
        //    //{
        //    //    double _val = nrRegulon + 1.0;
        //    //    chartArea.AxisY.CustomLabels.Add(_val - 0.25, _val + 0.25, String.Format("{0}", regulon), 0, LabelMarkStyle.None, GridTickTypes.TickMark);

        //    //    System.Windows.Forms.DataVisualization.Charting.Series aSerie = enrichmentChart.Series.Add(String.Format("{0}", regulon));
        //    //    aSerie.ChartType = SeriesChartType.ErrorBar;
        //    //    DataPointCollection dataPoints = aSerie.Points;

        //    //    for (int p = 0; p < fc[nrRegulon].Length; p++)
        //    //    {

        //    //        int val = nrRegulon;
        //    //        float _fc = fc[nrRegulon][p];
        //    //        {
        //    //            dataPoints.AddXY(_fc, val + 1, val + 1 - 0.25, val + 1 + 0.25);
        //    //        }
        //    //    }

        //    //    nrRegulon++;

        //    //}


        //    //enrichmentChart.Titles.Add("enrichment plot");

        //    //chartArea.AxisX.Title = "fold change";
        //    //chartArea.AxisY.Title = "regulon";

        //    ////chartArea.RecalculateAxesScale();

        //    //enrichmentChart.ChartAreas.Add(chartArea);
        //    //enrichmentChart.Palette = (ChartColorPalette)Properties.Settings.Default.defaultPalette;

        //    //return CreateEmfShape(enrichmentChart, hWnd, theApp);


        //}

        //public Excel.Shape DrawEnrichmentChart_OLD(HashSet<string> aRegulons, SysData.DataTable aTable)
        //{
        //    if (theApp == null)
        //        return null;

        //    IntPtr hWnd = (IntPtr) theApp.Hwnd;

        //    nrGenes = aTable.Rows.Count;
        //    nrRegulons = aRegulons.Count;

        //    List<float[]> fc = new List<float[]>();

        //    SysData.DataView dataView = aTable.AsDataView();

        //    double MMAX = (double)(float)aTable.Rows[0]["FC"];
        //    double MMIN = (double)(float)aTable.Rows[0]["FC"];


        //    foreach (string regulon in aRegulons)
        //    {
        //        dataView.RowFilter = String.Format("Regulon = '{0}'", regulon);
        //        SysData.DataTable dataTable = dataView.ToTable();
        //        int nrRows = dataTable.Rows.Count;
        //        float[] vs = new float[nrRows];
        //        for (int _r = 0; _r < nrRows; _r++)
        //        {
        //            double _val = (double)(float)dataTable.Rows[_r]["FC"];
        //            if (_val > MMAX) { MMAX = _val; }
        //            if (_val < MMIN) { MMIN = _val; }
        //            vs[_r] = (float)_val;

        //        }

        //        string toClipboard = string.Join("\n", vs.ToArray());
        //        System.Windows.Forms.Clipboard.SetData(System.Windows.Forms.DataFormats.Text, toClipboard);

        //        fc.Add(vs);
        //    }


            
        //    enrichmentChart = new Chart();
        //    enrichmentChart.Legends.Clear();
        //    ChartArea chartArea = new ChartArea("enrichmentChart");
            
        //    enrichmentChart.Height = nrRegulons*40; 
        //    enrichmentChart.Width = 400;

        //    chartArea.AxisX.MajorGrid.Enabled = false;
        //    chartArea.AxisY.MajorGrid.Enabled = false;
            

        //    chartArea.AxisY.Minimum = 0;
        //    chartArea.AxisY.Maximum = nrRegulons + 1;

        //    //enrichmentChart.ChartAreas.Add(chartArea);
            
        //    //float stepsize = CalculateStepSize((float)(MMAX - MMIN), 5);

        //    //double lMin = Math.Sign(MMIN) * Math.Ceiling(Math.Abs(MMIN));
        //    //double lMax = Math.Sign(MMAX) * Math.Ceiling(Math.Abs(MMAX));


        //    //for (int i = 0; i < 5; i++)
        //    //{
        //    //    chartArea.AxisX.CustomLabels.Add(lMin + (i * stepsize), lMin + ((i+1)* stepsize), String.Format("{0}", lMin + (i * stepsize)), 0, LabelMarkStyle.SideMark, GridTickTypes.TickMark);
        //    //}

        //    //chartArea.AxisX.Maximum = lMax;
        //    //chartArea.AxisX.Minimum = lMin;


        //    SetAxisTickLabels(chartArea.AxisX,MMIN,MMAX);


        //    int nrRegulon = 0;

        //    foreach (string regulon in aRegulons)
        //    {
        //        double _val = nrRegulon + 1.0;
        //        chartArea.AxisY.CustomLabels.Add(_val - 0.25, _val + 0.25, String.Format("{0}", regulon), 0, LabelMarkStyle.None, GridTickTypes.TickMark);

        //        System.Windows.Forms.DataVisualization.Charting.Series aSerie = enrichmentChart.Series.Add(String.Format("{0}", regulon));
        //        aSerie.ChartType = SeriesChartType.ErrorBar;
        //        DataPointCollection dataPoints = aSerie.Points;
                
        //        for (int p = 0; p < fc[nrRegulon].Length; p++)
        //        {

        //            int val = nrRegulon;
        //            float _fc = fc[nrRegulon][p];
        //            {
        //                dataPoints.AddXY(_fc, val + 1, val +1 - 0.25, val + 1 + 0.25);
        //            }                    
        //        }

        //        nrRegulon++;

        //    }

          
        //    enrichmentChart.Titles.Add("enrichment plot");

        //    chartArea.AxisX.Title = "fold change";
        //    chartArea.AxisY.Title = "regulon";

        //    //chartArea.RecalculateAxesScale();

        //    enrichmentChart.ChartAreas.Add(chartArea);
        //    enrichmentChart.Palette = (ChartColorPalette) Properties.Settings.Default.defaultPalette;

        //    return CreateEmfShape(enrichmentChart, hWnd, theApp);


        //}



        public static (float,float) CalculateStepSize(float range, float targetSteps)
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



    // from https://csharp.hotexamples.com/site/file?hash=0xe190e190f18b65d6b60ebe89ec697dff1cbe929ec830ef4f083e491a767f7c31&fullName=Source/EmfHelper.cs&project=LudovicT/NShape
    internal static class EmfHelper
    {
        #region Methods

        /// <summary>
        /// Copies the given <see cref="T:System.Drawing.Imaging.MetaFile" /> to the clipboard.
        /// The given <see cref="T:System.Drawing.Imaging.MetaFile" /> is set to an invalid state inside this function.
        /// </summary>
        public static bool PutEnhMetafileOnClipboard(IntPtr hWnd, Metafile metafile)
        {
            return PutEnhMetafileOnClipboard(hWnd, metafile, true);
        }

        /// <summary>
        /// Copies the given <see cref="T:System.Drawing.Imaging.MetaFile" /> to the clipboard.
        /// The given <see cref="T:System.Drawing.Imaging.MetaFile" /> is set to an invalid state inside this function.
        /// </summary>
        public static bool PutEnhMetafileOnClipboard(IntPtr hWnd, Metafile metafile, bool clearClipboard)
        {
            if (metafile == null) throw new ArgumentNullException("metafile");
            bool bResult = false;
            IntPtr hEMF, hEMF2;
            hEMF = metafile.GetHenhmetafile(); // invalidates mf
            if (!hEMF.Equals(IntPtr.Zero))
            {
                try
                {
                    hEMF2 = CopyEnhMetaFile(hEMF, null);
                    if (!hEMF2.Equals(IntPtr.Zero))
                    {
                        if (OpenClipboard(hWnd))
                        {
                            try
                            {
                                if (clearClipboard)
                                {
                                    if (!EmptyClipboard())
                                        return false;
                                }
                                IntPtr hRes = SetClipboardData(14 /*CF_ENHMETAFILE*/, hEMF2);
                                bResult = hRes.Equals(hEMF2);
                            }
                            finally
                            {
                                CloseClipboard();
                            }
                        }
                    }
                }
                finally
                {
                    DeleteEnhMetaFile(hEMF);
                }
            }
            return bResult;
        }

        /// <summary>
        /// Copies the given <see cref="T:System.Drawing.Imaging.MetaFile" /> to the specified file. If the file does not exist, it will be created.
        /// The given <see cref="T:System.Drawing.Imaging.MetaFile" /> is set to an invalid state inside this function.
        /// </summary>
        public static bool SaveEnhMetaFile(string fileName, Metafile metafile)
        {
            if (metafile == null) throw new ArgumentNullException("metafile");
            bool result = false;
            IntPtr hEmf = metafile.GetHenhmetafile();
            if (hEmf != IntPtr.Zero)
            {
                IntPtr resHEnh = CopyEnhMetaFile(hEmf, fileName);
                if (resHEnh != IntPtr.Zero)
                {
                    DeleteEnhMetaFile(resHEnh);
                    result = true;
                }
                DeleteEnhMetaFile(hEmf);
                metafile.Dispose();
            }
            return result;
        }

        [DllImport("user32.dll")]
        static extern bool CloseClipboard();

        [DllImport("gdi32.dll")]
        static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string fileName);

        [DllImport("gdi32.dll")]
        static extern bool DeleteEnhMetaFile(IntPtr hemf);

        [DllImport("user32.dll")]
        static extern bool EmptyClipboard();

        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

        #endregion Methods
    }
}
