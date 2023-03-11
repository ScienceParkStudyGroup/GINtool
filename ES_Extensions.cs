using System.Collections.Generic;
using System.Linq;
using System;
using System.Security.Cryptography;
using System.Text;
using Accord.Collections;
using Accord.Statistics.Distributions.Multivariate;
using Accord;
using System.ComponentModel;
using Accord.Math;
using Microsoft.Office.Interop.Excel;

namespace GINtool
{
    using stat_dict = System.Collections.Generic.Dictionary<string, double>;
    using rank_dict = System.Collections.Generic.Dictionary<string, int>;
    using dict_rank = System.Collections.Generic.Dictionary<int, string>;
    using lib_dict = System.Collections.Generic.Dictionary<string, string[]>;
    using LoessFunc = System.Func<int, double>;

    public static class ES_Extensions
    {
        // from https://stackoverflow.com/questions/4823467/using-linq-to-find-the-cumulative-sum-of-an-array-of-numbers-in-c-sharp
        public static IEnumerable<double> CumulativeSum(this IEnumerable<double> sequence)
        {
            double sum = 0;
            foreach (var item in sequence)
            {
                sum += item;
                yield return sum;
            }
        }

        // https://stackoverflow.com/questions/2733541/what-is-the-best-way-to-implement-this-composite-gethashcode
        public static int GetHashCodeValue(this IEnumerable<string> ar)
        {
            ar = ar.OrderBy(x => x);
            int hash = 3;
            unchecked
            {
                foreach (string s in ar)
                {
                    hash = hash * 5 + s.GetHashCode();
                }
                // Maybe nullity checks, if these are objects not primitives!                            
            }
            return hash;

        }
        public static IEnumerable<double> CumulativeMin(this IEnumerable<double> sequence)
        {
            double m = sequence.First();
            foreach (var item in sequence)
            {
                if (item < m)
                    m = item;

                yield return m;
            }
        }

        public static IEnumerable<int> CumulativeSum(this IEnumerable<int> sequence)
        {
            int sum = 0;
            foreach (var item in sequence)
            {
                sum += item;
                yield return sum;
            }
        }

        public static rank_dict RankMap(this stat_dict lib)
        {
            return (rank_dict)ParallelEnumerable.Range(0, lib.Count()).Select(i => i).ToDictionary(i => lib.ElementAt(i).Key, i => i);
        }

        public static dict_rank MapRank(this stat_dict lib)
        {
            return (dict_rank)ParallelEnumerable.Range(0, lib.Count()).Select(i => i).ToDictionary(i => i, i => lib.ElementAt(i).Key);
        }

        public static IEnumerable<rank_map> RankMap(this stat_map[] sequence)
        {
            // sequence.ToList().Sort((x, y) => x.Stat.CompareTo(y.Stat));

            sequence = sequence.OrderBy(p => p.Stat).Reverse().ToArray();
            int order = 0;
            foreach (var item in sequence)
            {
                yield return new rank_map(order++, item.Label);
            }
        }
        public static string[] subset(this stat_map[] s, int[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => s[y[i]].Label).ToArray();
        }

        public static double[] Plus(this double[] x, double[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => x[i] + y[i]).ToArray();
        }

        public static string[] subset(this stat_dict s, int[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => s.ElementAt(y[i]).Key).ToArray();
        }

        public static string[] subset2(this stat_dict s, string[] y)
        {
            return s.Keys.Intersect(y).ToArray();
            //return ParallelEnumerable.Range(0, matches.Count()).Select(i => s[matches[i]]).ToArray();

        }

        public static IEnumerable<double> Range(double min, double max, int nrpoints)
        {
            double step = (max - min) / nrpoints;
            double result = min;
            for (int i = 0; result < max; i++)
            {
                result = min + (step * i);
                yield return result;
            }
        }

        // https://www.statsmodels.org/stable/_modules/statsmodels/stats/multitest.html
        public static double Log1p(double x)
        {
            return Math.Log(x + 1);
        }

        public static double Exp1m(double x)
        {
            return Math.Exp(x) - 1;
        }


        public static double[] sidak_correction(double[] pvals)
        {

            /*  reject = pvals <= alphacSidak
        pvals_corrected = -np.expm1(ntests * np.log1p(-pvals))
             */

            int ntests = pvals.Length;
            return pvals.Select(x => -Exp1m(ntests * Log1p(-x))).ToArray();
        }

        public static double[] ecdf(int n)
        {
            // assume sorted array

            return Enumerable.Range(1, n).Select(x => (double)x / (double)n).ToArray();

            // ys = np.cumsum(1 for _ in x)/ float(len(x))
        }

        public static double[] fdr_correction(double[] pvals, double alpha = 0.05)
        {
            // https://stackoverflow.com/questions/10443461/c-sharp-array-findallindexof-which-findall-indexof
            List<int> _missings = Enumerable.Range(0, pvals.Length).Where(i => double.IsNaN(pvals[i])).ToList();
            //pvals = pvals.Where(x => !double.IsNaN(x)).ToArray();
            var sortedElements = pvals.Select((x, i) => new KeyValuePair<double, int>(x, i)).OrderBy(x => x.Key).ToArray();

            List<int> sortedIndex = sortedElements.Select(x => x.Value).ToList();

            double[] pvals_sorted = Enumerable.Range(0, sortedIndex.Count).Select(i => pvals[sortedIndex[i]]).ToArray();

            // double[] pvals_sorted = pvals.OrderBy(x => x).ToArray();
            double[] ecdffactor = ecdf(pvals_sorted.Length);
            double[] pvals_corrected_raw = Enumerable.Range(0, pvals.Length).Select(i => pvals_sorted[i] / ecdffactor[i]).ToArray();
            double[] pvals_corrected = pvals_corrected_raw.Reverse().CumulativeMin().Reverse().ToArray();
            for (int i = 0; i < pvals_corrected.Length; i++)
            {
                if (pvals_corrected[i] > 1)
                    pvals_corrected[i] = 1;
            }
            pvals_corrected_raw = Enumerable.Repeat(0.0, pvals_corrected.Length).ToArray();

            for (int i = 0; i < pvals_corrected.Length; i++)
            {
                pvals_corrected_raw[sortedIndex[i]] = pvals_corrected[i];
            }

            foreach (int i in _missings)
                pvals_corrected_raw[i] = double.NaN;

            return pvals_corrected_raw;

        }


        public static double percentile(int[] sortedData, double p)
        {
            // algo derived from Aczel pg 15 bottom
            if (p >= 100.0d) return sortedData[sortedData.Length - 1];

            double position = (double)(sortedData.Length + 1) * p / 100.0;
            double leftNumber = 0.0d, rightNumber = 0.0d;

            double n = p / 100.0d * (sortedData.Length - 1) + 1.0d;

            if (position >= 1)
            {
                leftNumber = sortedData[(int)System.Math.Floor(n) - 1];
                rightNumber = sortedData[(int)System.Math.Floor(n)];
            }
            else
            {
                leftNumber = sortedData[0]; // first data
                rightNumber = sortedData[1]; // first data
            }

            if (leftNumber == rightNumber)
                return leftNumber;
            else
            {
                double part = n - System.Math.Floor(n);
                return leftNumber + part * (rightNumber - leftNumber);
            }
        }

        public static IEnumerable<int> toint(this double[] array)
        {
            return ParallelEnumerable.Range(0, array.Length).Select(pp => Convert.ToInt32(array[pp]));
        }
        public static IEnumerable<double> percentiles(int[] array, double[] p)
        {
            return ParallelEnumerable.Range(0, p.Length).Select(pp => percentile(array, p[pp]));
        }

        public static string[] keys(this libitem[] s)
        {
            return ParallelEnumerable.Range(0, s.Length).Select(i => s[i].Label).ToArray();
        }
        public static int[] lengths(this libitem[] s)
        {
            return ParallelEnumerable.Range(0, s.Length).Select(i => s[i].Items.Length).ToArray();
        }

        public static int[] lengths(this lib_dict s)
        {
            return s.AsParallel().Select(x => x.Value.Length).ToArray();
        }

        public static double[] AbsStats(this stat_map[] stats)
        {
            return stats.Select(st => Math.Abs(st.Stat)).ToArray();
        }

        public static double[] AbsStats(this stat_dict stats)
        {
            return stats.Select(k => Math.Abs(k.Value)).ToArray();
        }


        public static double Sum(this double[] arr, int[] index)
        {
            double sum = 0;
            foreach (int i in index)
                sum += arr[i];
            return sum;
        }

        public static byte[] Hashvalue(Object x)
        {
            var tmpSource = ASCIIEncoding.ASCII.GetBytes(x.ToString());
            return new MD5CryptoServiceProvider().ComputeHash(tmpSource);

        }

        #region Linq operators
        // from https://stackoverflow.com/questions/59012299/multiplying-arrays-element-wise-has-unexpected-performance-in-c-sharp

        public static int[] Ranks(rank_map[] signatures)
        {
            return signatures.Select(signature => signature.Rank).ToArray();
        }

        public static double[] Pmult(double[] x, double[] y)
        {
            return ParallelEnumerable.Range(0, x.Length).Select(i => x[i] * y[i]).ToArray();
        }

        public static double[] Pmult(double[] x, int[] y)
        {
            return ParallelEnumerable.Range(0, x.Length).Select(i => x[i] * y[i]).ToArray();
        }
        public static int[] Pmult(rank_map[] x, int[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => x[i] * y[i]).ToArray();
        }

        public static int[] Pmult(int[] x, int[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => x[i] * y[i]).ToArray();
        }

        public static double[] Pmult(int[] y, double x)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => x * (double)y[i]).ToArray();
        }

        public static double[] Pmult(double[] y, double x)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => x * y[i]).ToArray();
        }

        public static double[] Pmult(double[] y, int x)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => x * y[i]).ToArray();
        }

        public static double[] Pmin(double[] y, int[] x)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => y[i] - (double)x[i]).ToArray();
        }

        public static double[] Pmin(double[] y, double[] x)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => y[i] - x[i]).ToArray();
        }

        public static double[] Pabs(this double[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => Math.Abs(y[i])).ToArray();
        }

        //public static double[] Pabs(double[] y)
        //{
        //    return ParallelEnumerable.Range(0, y.Length).Select(i => Math.Abs(y[i])).ToArray();
        //}

        public static string[] subset(this string[] s, int[] y)
        {
            return ParallelEnumerable.Range(0, y.Length).Select(i => s[y[i]]).ToArray();
        }


        //public static List<string> strip_gene_set(List<string> signature_genes, List<string> gene_set)
        //{
        //    return gene_set.Intersect(signature_genes).ToList();

        //}

        public static string Join(this stat_map[] signatures, int[] subset)
        {
            return string.Join(",", ParallelEnumerable.Range(0, subset.Length).Select(i => signatures[subset[i]].Label).ToArray());

        }

        #endregion

     
    };



}