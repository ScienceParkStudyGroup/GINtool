using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GINtool
{
    public static class ClassExtensions
    {
        public static DateTime Tomorrow(this DateTime date)
        {
            return date.AddDays(1);
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

        public static float sd(this List<float> list)
        {
            if (list.Count == 1)
                return 0;

            double avg = list.Average();
            double sd = 0;
            for (int i = 0; i < list.Count; i++)
            {
                sd += Math.Pow(list[i] - avg, 2.0);
            }

            return (float)Math.Sqrt(sd / (list.Count - 1));
        }

        public static float median(this List<float> list)
        {
            int n = list.Count;

            list.Sort();
            int m = n / 2; // +1 -1 (for zero based counting)
            int r = n % 2;
            if (r == 0 && n > 1) // if even number of elements            
                return (list[m - 1] + list[m]) / 2;
            return list[m];
        }

        public static float mad(this List<float> list)
        {
            if (list.Count == 1)
                return 0;

            float median = list.median();
            List<float> md = new List<float>(list.Count);
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

    }
}
