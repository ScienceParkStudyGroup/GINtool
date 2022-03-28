using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;

namespace GINtool
{
    internal static class ClassExtensions
    {

        public static int ParseInt(string str, int defaultVal = 0)
        {
            int _val; bool _res = int.TryParse(str, out _val);
            if (_res)
                return _val;
            else
                return defaultVal;
        }

        public static DateTime Tomorrow(this DateTime date)
        {
            return date.AddDays(1);
        }

        public static double fisherz(double r)
        {
            // this should not happen but sometimes if an error was made the p value is maximized to 1 by default.
            if (r > 1)
                r = 0.99999;
            // this is an approximation of course but otherwise inf and nan issues
            if (Math.Abs(r - 1) < Double.Epsilon)
                r = 0.99999; // 1 - Double.Epsilon; 
            return 0.5 * Math.Log((1 + r) / (1 - r));
        }

        public static double fzBack(double avgZ_h)
        {
            return (Math.Exp(avgZ_h) - Math.Exp(-avgZ_h)) / (Math.Exp(avgZ_h) + Math.Exp(-avgZ_h));
        }


        public static double Gauss(double z)
        {
            // input = z-value (-inf to +inf)
            // output = p under Standard Normal curve from -inf to z
            // e.g., if z = 0.0, function returns 0.5000
            // ACM Algorithm #209
            double y; // 209 scratch variable
            double p; // result. called 'z' in 209
            double w; // 209 scratch variable
            if (z == 0.0)
                p = 0.0;
            else
            {
                y = Math.Abs(z) / 2;
                if (y >= 3.0)
                {
                    p = 1.0;
                }
                else if (y < 1.0)
                {
                    w = y * y;
                    p = ((((((((0.000124818987 * w
                      - 0.001075204047) * w + 0.005198775019) * w
                      - 0.019198292004) * w + 0.059054035642) * w
                      - 0.151968751364) * w + 0.319152932694) * w
                      - 0.531923007300) * w + 0.797884560593) * y * 2.0;
                }
                else
                {
                    y = y - 2.0;
                    p = (((((((((((((-0.000045255659 * y
                      + 0.000152529290) * y - 0.000019538132) * y
                      - 0.000676904986) * y + 0.001390604284) * y
                      - 0.000794620820) * y - 0.002034254874) * y
                      + 0.006549791214) * y - 0.010557625006) * y
                      + 0.011630447319) * y - 0.009279453341) * y
                      + 0.005353579108) * y - 0.002141268741) * y
                      + 0.000535310849) * y + 0.999936657524;
                }
            }
            if (z > 0.0)
                return (p + 1.0) / 2;
            else
                return (1.0 - p) / 2;
        }


        public static double Student(double t, double df)
        {
            // for large integer df or double df
            // adapted from ACM algorithm 395
            // returns 2-tail p-value
            double n = df; // to sync with ACM parameter name
            double a, b, y;
            t = t * t;
            y = t / n;
            b = y + 1.0;
            if (y > 1.0E-6) y = Math.Log(b);
            a = n - 0.5;
            b = 48.0 * a * a;
            y = a * y;
            y = (((((-0.4 * y - 3.3) * y - 24.0) * y - 85.5) /
              (0.8 * y * y + 100.0 + b) + y + 3.0) / b + 1.0) *
              Math.Sqrt(y);
            return 2.0 * Gauss(-y); // ACM algorithm 209
        }


        public static double paverage(this List<double> list)
        {
            List<double> _fz = list.Select(x => fisherz(x)).ToList();
            double AvgP = _fz.Average();
            return fzBack(AvgP);
        }

        public static double sd(this List<double> list)
        {
            if (list.Count == 1)
                return 0;

            double avg = list.Average();
            double sd = 0;
            for (int i = 0; i < list.Count; i++)
            {
                sd += Math.Pow(list[i] - avg, 2.0);
            }

            return Math.Sqrt(sd / (list.Count - 1));
        }

        public static double median(this List<double> list)
        {
            int n = list.Count;

            list.Sort();
            int m = n / 2; // +1 -1 (for zero based counting)
            int r = n % 2;
            if (r == 0 && n > 1) // if even number of elements            
                return (list[m - 1] + list[m]) / 2;
            return list[m];
        }

        public static double mad(this List<double> list)
        {
            if (list.Count == 1)
                return 0;

            double median = list.median();
            List<double> md = new List<double>(list.Count);
            for (int i = 0; i < list.Count; i++)
                md.Add(Math.Abs(list[i] - median));

            return md.median();

        }

        public static double AbsMad(this List<double> list)
        {
            if (list.Count == 1)
                return 0;

            list = list.Select(x => Math.Abs(x)).ToList();
            double median = list.median();
            List<double> md = new List<double>(list.Count);
            for (int i = 0; i < list.Count; i++)
                md.Add(Math.Abs(list[i] - median));

            return md.median();

        }

        public static bool Any(this byte en)
        {
            return ((byte)GinRibbon.UPDATE_FLAGS.NONE | en) != (byte)GinRibbon.UPDATE_FLAGS.NONE;
        }

        public static bool None(this byte en)
        {
            return en == (byte)GinRibbon.UPDATE_FLAGS.NONE;
        }

        public static bool Check(this byte en, GinRibbon.UPDATE_FLAGS val)
        {
            return (en & (byte)val) == (byte)val;
        }

        public static bool GetBit(this byte b, int bitNumber)
        {
            return (b & (1 << bitNumber)) != 0;
        }


        public static summaryInfo GetCatValues(this List<summaryInfo> sis, string catName)
        {
            if (sis != null)
                return sis.Where(x => x.catName == catName).ToArray()[0];

            return new summaryInfo();
        }

        public static cat_elements GetCatElement(this List<cat_elements> el, string catName)
        {
            if (el != null && el.Count > 0)
            {
                if (el != null)
                {
                    IEnumerable<cat_elements> output = el.Where(x => x.catName == catName);
                    if (output.Count() > 0)
                    {
                        return output.First();
                    }

                }
            }
            return new cat_elements();
        }

        //public static int FindAssocForRegulon(this BsuLinkedItems bsu, string regulon)
        //{
        //    int _pos = bsu.Regulons.IndexOf(regulon);
        //    if(bsu.REGULON_UP.Contains(_pos))
        //    {
        //        return _pos;
        //    }
        //    if(bsu.REGULON_DOWN.Contains(_pos))
        //    {
        //        return -_pos;
        //    }
        //    return 0;
        //}

        public static BsuLinkedItems GetByGeneName(this List<BsuLinkedItems> lst, string geneName)
        {
            if (lst != null && lst.Count > 0)
            {
                IEnumerable<BsuLinkedItems> output = lst.Where(x => x.GeneName == geneName);
                if (output.Count() > 0)
                {
                    return output.First();
                }
            }

            return new BsuLinkedItems();
        }

        public static double AbsAverage(this List<double> lst)
        {
            if (lst.Count > 0)
            {
                return lst.Select(x => Math.Abs(x)).ToArray().Average();
            }
            return Double.NaN;
        }


#if CLICK_CHART
        public static string getPoint(this chart_info chart, int serie, int point)
        {
            return "";
        }


        public static readonly chart_info Empty = new chart_info();


        public static chart_info isFound(this List<chart_info> chart_Infos, object chart)
        {
           
            foreach(chart_info cinfo in chart_Infos)
            {
                if (cinfo.chart.Equals(chart))
                    return cinfo;                                
            }
            return Empty;
        }
#endif

        //}



        // from https://csharp.hotexamples.com/site/file?hash=0xe190e190f18b65d6b60ebe89ec697dff1cbe929ec830ef4f083e491a767f7c31&fullName=Source/EmfHelper.cs&project=LudovicT/NShape
        //internal static class EmfHelper
        //{
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
