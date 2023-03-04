using System.Security.Cryptography;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System;
using System.Runtime.CompilerServices;
using System.Collections.Specialized;
using System.Dynamic;
using System.Security.Policy;
using System.Text;
using System.Runtime.Serialization;
using Accord.IO;

namespace GINtool
{

    //public class stat_dict : Dictionary<string, double> { }
    //public class rank_dict : Dictionary<string, int> { }
    using stat_dict = Dictionary<string, double>;
    using rank_dict = Dictionary<string, int>;
    //using LoessFunc = Func<int, double>;

    [Serializable]
    public class LoessFunc
    {
        //Func<int, double>, 
        //public int set_size;
        int size;
        double[] _xvalues;
        double[] _yvalues;
        double[] _ypredict;

        public LoessFunc()
        {
            size = 0;
        }

        public LoessFunc(double[] xvalues, double[] yvalues)
        {
            _xvalues = xvalues;
            _yvalues = yvalues;
            size = _xvalues.Length;
        }

        public void fit(int[] xvalues, double[] yvalues, double frac = 0.5)
        {
            _xvalues = xvalues.Select(Convert.ToDouble).ToArray();
            _yvalues = yvalues;
            size = _yvalues.Length;
            frac = Math.Max(frac, 2 / (double)xvalues.Length);
            LoessInterpolator loess = new LoessInterpolator(bandwidth: frac, robustnessIters: 2);
            _ypredict = loess.smooth(_xvalues, _yvalues);

            _ypredict = _ypredict.Select(i => Double.IsNaN(i) ? 0 : (Double.IsInfinity(i) ? 1 : i)).ToArray();
        }

        public void fit(double[] xvalues, double[] yvalues, double frac = 0.5)
        {
            _xvalues = xvalues;
            _yvalues = yvalues;
            size = yvalues.Length;
            frac = Math.Max(frac, 2 / (double)xvalues.Length);
            LoessInterpolator loess = new LoessInterpolator(bandwidth: frac, robustnessIters: 2);
            _ypredict = loess.smooth(_xvalues, _yvalues);
            _ypredict = _ypredict.Select(i => Double.IsNaN(i) ? 0 : (Double.IsInfinity(i) ? 1 : i)).ToArray();
        }


        public double predict(int x)
        {
            if (_xvalues.Length == 0)
                return x;

            int idx1 = Enumerable.Range(0, _xvalues.Length).Where(y => _xvalues[y] < x).Count() - 1;

            if (idx1 == _xvalues.Length)
                return _ypredict.Last();
            if (idx1 <= 0)
                return _ypredict.First();

            double xrange = _xvalues[idx1 + 1] - _xvalues[idx1];
            double yrange = (_ypredict[idx1 + 1] - _ypredict[idx1]);
            double dp = (x - _xvalues[idx1]) / xrange;
            return dp * yrange + _ypredict[idx1];

            //return -1.0;
        }
    }

    public class rank_map
    {
        public int Rank;
        public string Label;
        //public int Rank { get => rank; set => rank = value; }
        //public string Label { get => label; set => label = value; }
        public rank_map(int r, string l)
        {
            Rank = r;
            Label = l;
        }
        public rank_map(string l, int r)
        {
            Rank = r;
            Label = l;
        }
        public static int operator *(rank_map a, int b)
        {
            return a.Rank * b;
        }

        //public static int[] Ranks(this rank_map[] map,string[] selection)
        //{
        //    //map.Where(i=>selection.Co(i)).
        //}


    }


    public class stat_map
    {
        double stat;
        string label;
        public double Stat { get => stat; set => stat = value; }
        public string Label { get => label; set => label = value; }
        public stat_map() { }
        public stat_map(double r, string l)
        {
            stat = r;
            label = l;
        }
        public stat_map(string l, double r)
        {
            stat = r;
            label = l;
        }


        public static stat_map FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(',');
            stat_map map = new stat_map();
            map.Label = values[0];
            map.Stat = Convert.ToDouble(values[1]);
            return map;
        }

    }


    public class libitem
    {
        string label;
        string[] items;
        public string[] Items { get => items; set => items = value; }
        public string Label { get => label; set => label = value; }
        public libitem(string l, string[] its)
        {
            label = l; items = its;
        }
        public libitem(string[] its, string l)
        {
            label = l; items = its;
        }
    }

    public struct S_GSEA
    {
        // public string label;
        public double pval, sidak, es, nes, fdr;
        public int size;
        public string leading_edge;
    }

    [Serializable]
    public class S_ESPARAMS
    {
        public LoessFunc alpha_pos, beta_pos, pos_ratio;
        public double ks_pos, ks_neg;

        public S_ESPARAMS()
        { }
        
    }
}