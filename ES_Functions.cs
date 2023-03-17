using System.Collections.Generic;
using System.Linq;
using System;
using System.Diagnostics;
using Accord.Statistics.Distributions.Univariate;
using Accord.Statistics.Testing;
using Accord.Math;
using System.Collections;
using System.Threading;
using System.Threading.Tasks;
using static GINtool.ES_Extensions;
using System.Data;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace GINtool
{

    using stat_dict = Dictionary<string, double>;
    using rank_dict = Dictionary<string, int>;
    using dict_rank = Dictionary<int, string>;
    using gsea_dict = Dictionary<string, S_GSEA>;
    using lib_dict = Dictionary<string, string[]>;
    using dataset_dict = Dictionary<string, GINtool.DataItem>;

    public static class ES_Functions
    {

        public static (double[], double) enrichment_score(double[] abs_signature, rank_dict signature_map, string[] gene_set)
        {
            //int[] hits = gene_set.Select(item => signature_map.ToList().Find(s => s.Label.Equals(item)).Rank).ToArray();            
            string[] matches = signature_map.Keys.Intersect(gene_set).ToArray();
            int[] hits = ParallelEnumerable.Range(0, matches.Count()).Select(i => signature_map[matches[i]]).ToArray();

            int[] hit_indicator = Enumerable.Repeat(0, abs_signature.Length).ToArray();
            int[] no_hit_indicator = Enumerable.Repeat(1, abs_signature.Length).ToArray();
            foreach (int r in hits)
            {
                hit_indicator[r] = 1;
                no_hit_indicator[r] = 0;
            }

            int number_hits = hits.Count();
            int number_miss = abs_signature.Length - number_hits;
            double sum_hit_scores = abs_signature.Sum(hits);
            double[] emptyResult = { 0.0 };

            if (sum_hit_scores == 0)
                return (emptyResult, 0);

            double norm_hit = 1.0 / sum_hit_scores;
            double norm_no_hit = 1.0 / (double)number_miss;

            double[] rsum_neg = Pmult(no_hit_indicator, norm_no_hit);
            double[] rsum_pos = Pmult(Pmult(abs_signature, hit_indicator), norm_hit);
            double[] running_sum = Pmin(rsum_pos, rsum_neg).CumulativeSum().ToArray();

            double[] abs_rs = running_sum.Pabs();
            int indexAtMax = abs_rs.ToList().IndexOf(abs_rs.Max());
            double es = running_sum[indexAtMax];

            return (running_sum, es);
        }

        public static string[] strip_gene_set(string[] signature_genes, string[] gene_set)
        {
            return gene_set.Intersect(signature_genes).ToArray();

        }

        public static string get_leading_edge(double[] runningsum, dict_rank map_signature, string[] gene_set, rank_dict signature_map)
        {
            // van genes naar ranks
            int[] hits = gene_set.Select(g => signature_map[g]).ToArray();
            //gene_set.Select(item => signature_map.ToList().Find(s => s.Label.Equals(item)).Rank).ToArray();
            int rmax = runningsum.ToList().IndexOf(runningsum.Max());
            int rmin = runningsum.ToList().IndexOf(runningsum.Min());
            int[] lgenes;
            // kijk of het negatief of juist positief is
            if (runningsum[rmax] > Math.Abs(runningsum[rmin]))
            {
                lgenes = hits.Intersect(Enumerable.Range(0, rmax)).ToArray();
            }
            else
            {
                lgenes = hits.Intersect(Enumerable.Range(rmin, runningsum.Length)).ToArray();
            }
            // ga weer terug naar een set genen uit de originele lijst via lgenes (=positie van genen in de lijst)
            return String.Join(",", ParallelEnumerable.Range(0, lgenes.Length).Select(i => map_signature[lgenes[i]]).ToArray());

        }



        public static double[] get_peak_size(stat_dict signature, double[] abs_signature, rank_dict signature_map, int size, int permutations, int seed)
        {
            double[] es = new double[permutations];
            int sig_count = signature.Count();
            Accord.Math.Random.Generator.Seed = seed; //SEED
            var all_items = Vector.Create(signature_map.Keys.ToArray());

            int processorCount = Environment.ProcessorCount;
           
            Parallel.For(0, permutations, new ParallelOptions { MaxDegreeOfParallelism = processorCount }, i =>
            {
                (_, es[i]) = enrichment_score(abs_signature, signature_map, all_items.Sample(size));
            });


            return es;
        }

        public static LoessFunc Loess2(int[] x, double[] y, double frac = 0.5)
        {
            double[] _x = x.Select(Convert.ToDouble).ToArray();
            frac = Math.Max(frac, 2 / (double)x.Length);
            LoessFunc loess = new LoessFunc();
            loess.fit(x, y);
            //LoessInterpolator loess = new LoessInterpolator(bandwidth: frac, robustnessIters: 2);
            return loess;
        }

        public static LoessFunc loess_interpolation(double[] x, double[] y, double frac = 0.5)
        {
            LoessFunc loess = new LoessFunc();

            // LoessInterpolator loess = new LoessInterpolator(bandwidth: frac, robustnessIters: 2);
            loess.fit(x, y, frac);
            return loess;
        }

        public static LoessFunc loess_interpolation(int[] x, double[] y, double frac = 0.5)
        {
            double[] _x = x.Select(Convert.ToDouble).ToArray();
            frac = Math.Max(frac, 2 / (double)x.Length);
            LoessFunc loess = new LoessFunc();
            //LoessInterpolator loess = new LoessInterpolator(bandwidth: frac, robustnessIters: 2);
            loess.fit(_x, y, frac);
            return loess;
        }

        public static (double, double, double, double, double, double, double) estimate_anchor(stat_dict signature, double[] abs_signature, rank_dict signature_map, int set_size, int permutations, bool symmetric, int seed)
        {
         //   Stopwatch sw = Stopwatch.StartNew();
           // sw.Start();
            double[] es = get_peak_size(signature, abs_signature, signature_map, set_size, permutations, seed);
            //sw.Stop();

            //Console.WriteLine("normal thread " + sw.Elapsed.ToString() + " size:" + set_size.ToString());   

            //get_peak_size_thread(signature, abs_signature, signature_map, set_size, permutations, seed);

            double[] pos = es.Where(x => x > 0).ToArray();
            double[] neg = es.Where(x => x < 0).ToArray();

            double ks_pos, ks_neg, alpha_pos, beta_pos, alpha_neg, beta_neg;

            if ((neg.Length < 250 | pos.Length < 250) & (symmetric == false))
            {
                symmetric = true;
            }
            if (symmetric)
            {
                double[] aes = es.Pabs();

                GammaDistribution dist = new GammaDistribution();
                dist.Fit(aes);
                alpha_pos = dist.Shape;
                beta_pos = dist.Scale;

                var kstest = new KolmogorovSmirnovTest(aes, dist);
                ks_pos = kstest.PValue;
                ks_neg = kstest.PValue;

                alpha_neg = alpha_pos;
                beta_neg = beta_pos;
            }
            else
            {
                GammaDistribution dist = new GammaDistribution();
                dist.Fit(pos);
                alpha_pos = dist.Shape;
                beta_pos = dist.Scale;
                var kstest = new KolmogorovSmirnovTest(pos, dist);
                ks_pos = kstest.PValue;

                dist = new GammaDistribution();
                dist.Fit(neg.Pabs());
                alpha_neg = dist.Shape;
                beta_neg = dist.Scale;
                kstest = new KolmogorovSmirnovTest(neg.Pabs(), dist);
                ks_neg = kstest.PValue;
            }

            double pos_ratio = (double)pos.Length / (double)(pos.Length + neg.Length);

            return (alpha_pos, beta_pos, ks_pos, alpha_neg, beta_neg, ks_neg, pos_ratio);
        }


        public static S_ESPARAMS estimate_parameters(stat_dict signature, double[] abs_signatures, rank_dict signature_map, lib_dict library, int permutations = 2000, bool symmetric = false, int calibration_anchors = 20, int seed = 0)
        {
            int[] ll = library.lengths();
            ll.Sort(); // for percentile calculation
            
            //Count().lengths().OrderBy(v => v).ToArray();         
            // from https://stackoverflow.com/questions/1139181/a-method-to-count-occurrences-in-a-list

            int signature_count = signature.Count();

            //int[] cumsum = g.Select(grp => grp.Count()).CumulativeSum().ToArray();
            double[] q = Range(2, 100, calibration_anchors - 1).ToArray();
            // partion the library sizes in 5% groups .. this makes this method a-priori dependent on selection of genesets .. 
            double[] nn = percentiles(ll, q).OrderBy(v => v).ToArray();
            //double[] nn = q.Select(v => v / 100 * (signature_count - 1)).ToArray();

            
            int[] x = { 1, 4, 6, ll.Max(), signature_count / 2, signature_count - 1 };
            int[] y = nn.ToList().ConvertAll(Convert.ToInt32).ToArray();

            var z = new int[x.Length + y.Length];
            x.CopyTo(z, 0);
            y.CopyTo(z, x.Length);


            x = z.Distinct().OrderBy(i => i).ToArray();
            x = x.Where(i => i <= signature_count & i>0).ToArray(); // only in cases with very few items          

            double[] alpha_pos = new double[x.Length];
            double[] beta_pos = new double[x.Length];
            double[] ks_pos = new double[x.Length];

            double[] alpha_neg = new double[x.Length];
            double[] beta_neg = new double[x.Length];
            double[] ks_neg = new double[x.Length];

            double[] pos_ratio = new double[x.Length];

           // Stopwatch timer = new Stopwatch();
          //  timer.Start();

            int cnt = 0;
            foreach (int perc in x)
            {
                (alpha_pos[cnt], beta_pos[cnt], ks_pos[cnt], alpha_neg[cnt], beta_neg[cnt], ks_neg[cnt], pos_ratio[cnt]) =
                    estimate_anchor(signature, abs_signatures, signature_map, perc, permutations, symmetric, seed + perc);
                cnt++;
            }

            if (pos_ratio.Max() > 1.5)
            {
                MessageBox.Show("Significant unbalance between positive and negative enrichment scores detected. Signature values are not centered close to 0.");
            }
           // timer.Stop();
           // TimeSpan ts = timer.Elapsed;

          //  string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
          //  Console.WriteLine("RunTime " + elapsedTime);

            LoessFunc f_alpha_pos = loess_interpolation(x, alpha_pos);
            LoessFunc f_beta_pos = loess_interpolation(x, beta_pos, frac: 0.2);

            NormalDistribution normal = new NormalDistribution();
            double[] randn = normal.Generate(pos_ratio.Length);
            pos_ratio = Pmin(pos_ratio, randn.Multiply(0.0001).Pabs());
            LoessFunc f_pos_ratio = loess_interpolation(x, pos_ratio);

            S_ESPARAMS result = new S_ESPARAMS
            {
                alpha_pos = f_alpha_pos,
                beta_pos = f_beta_pos,
                pos_ratio = f_pos_ratio,
                ks_pos = ks_pos.Average(),
                ks_neg = ks_neg.Average()
            };

            return result;
        }


        public static void gsea_calibrate(dataset_dict dataset, lib_dict library, ref Hashtable hashtable, int permutations = 2000, int anchors = 20, bool verbose = false, bool symmetric = true, int seed = 0, bool pvalues=false)
        {
            if (permutations < 1000 && !symmetric)
            {
                if (verbose)
                    Console.WriteLine("Low numer of permutations can lead to inaccurate p-value estimation. Symmetric Gamma distribution enabled to increase accuracy");
                symmetric = true;
            }
            else if (permutations < 500)
            {
                if (verbose)
                    Console.WriteLine("Low numer of permutations can lead to inaccurate p-value estimation. Consider increasing number of permutations.");
                symmetric = true;
            }

            //stat_dict signature = dataset.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => Math.Abs(kvp.Value.FC));
            //stat_dict signature = dataset.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => kvp.Value.FC*-Math.Abs();
            stat_dict signature = getSignature(dataset, pvalues);


            // dataset.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => Math.Sign(kvp.Value.FC) * -Math.Log10(kvp.Value.pval));

            // stat_dict signature = dataset.ToDictionary(kvp => kvp.Key, kvp => Math.Abs(kvp.Value.FC));

            Random random = new Random(seed);

            stat_dict signature_ordered = signature.OrderBy(kvp => kvp.Value).ToDictionary(x => x.Key, x => x.Value);
            int sighashK = signature_ordered.Keys.GetHashCodeValue();
            int sighashV = signature_ordered.Values.GetHashCodeValue();
            int sig_hash = sighashK + sighashV;

            //byte[] sig_hash = Hashvalue(signature);
            //string sig_hash_str = System.Text.Encoding.Default.GetString(sig_hash);


            signature = signature.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
            rank_dict signature_map = signature.RankMap();
            dict_rank map_signature = signature.MapRank();


            NormalDistribution norm = new NormalDistribution();
            double[] sigvalues = signature.Values.ToArray();
            sigvalues = sigvalues.Plus(Pmult(norm.Generate(sigvalues.Length), 1 / (sigvalues.Average() * 10000)));
            double[] abs_signature = sigvalues.Abs();

            S_ESPARAMS es_params;
            if (!hashtable.ContainsKey(sig_hash))
            {
                if (verbose)
                    Console.WriteLine("Calibrating es parameters");
                es_params = estimate_parameters(signature, abs_signature, signature_map, library, permutations,calibration_anchors:anchors, seed:seed);
                hashtable.Add(sig_hash, es_params);
            }
            else
            {
                if (verbose)
                    Console.WriteLine("Loading previously calibrated results");
                es_params = (S_ESPARAMS)hashtable[sig_hash];
            }


        }

        public static S_GSEA gsea_calc_es(double[] abs_signature, dict_rank map_signature, rank_dict signature_map, S_ESPARAMS es_params, string[] geneset)
        {
            S_GSEA lib_result = new S_GSEA();
            int gsize = geneset.Length;
            double[] rs;
            double es;
            (rs, es) = enrichment_score(abs_signature, signature_map, geneset);
            string legenes = get_leading_edge(rs, map_signature, geneset, signature_map);

            double pos_alpha = es_params.alpha_pos.predict(gsize);
            double pos_beta = es_params.beta_pos.predict(gsize);
            double pos_ratio = es_params.pos_ratio.predict(gsize);

            GammaDistribution gamma = new GammaDistribution(pos_beta, pos_alpha);
            NormalDistribution normal = new NormalDistribution();
            double nes = 0;
            double prob_two_tailed;

            double prob;
            if (es > 0)
            {
                prob = 1 - gamma.ComplementaryDistributionFunction(es);
                prob_two_tailed = Math.Min(0.5, (1 - Math.Min((1 - pos_ratio) + prob * pos_ratio, 1)));

                if (prob_two_tailed < 1)
                {
                    nes = normal.InverseDistributionFunction(1 - Math.Min(1, prob_two_tailed));
                }

            }
            else
            {
                prob = 1 - gamma.ComplementaryDistributionFunction(-es);
                prob_two_tailed = Math.Min(0.5, (1 - Math.Min(prob * (1 - pos_ratio) + pos_ratio, 1)));
                nes = normal.InverseDistributionFunction(Math.Min(1, prob_two_tailed));
            }

            double pval = 2 * prob_two_tailed;
            lib_result.pval = pval;
            lib_result.es = es;
            lib_result.nes = nes;
            lib_result.size = gsize;
            lib_result.leading_edge = legenes;

            return lib_result;

        }

        public static string getHashValue(object x)
        {
            byte[] sig_hash = Hashvalue(x);
            return System.Text.Encoding.Default.GetString(sig_hash);
            
        }


        public static stat_dict getSignature(dataset_dict dataset, bool pvalues = false)
        {
            if (pvalues)
                return dataset.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => Math.Sign(kvp.Value.FC) * -Math.Log10(kvp.Value.pval));
            else
                return dataset.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => kvp.Value.FC);

        }


        // pre calculate expected genesets 
        public static void gsea_enrich(dataset_dict dataset, lib_dict library, Hashtable hashtable, ref Hashtable hashgsea, int min_size = 5, int max_size = 25000, bool pvalues=false)
        {
            //stat_dict signature = dataset.Where(kvp => kvp.Value.FC != 0).ToDictionary(kvp => kvp.Key, kvp => kvp.Value.FC);
            stat_dict signature = getSignature(dataset,pvalues);

            // string sig_hash_str = getHashValue(signature);
            stat_dict signature_ordered = signature.OrderBy(kvp => kvp.Value).ToDictionary(x => x.Key, x => x.Value);

            int sighashK = signature_ordered.Keys.GetHashCodeValue();
            int sighashV = signature_ordered.Values.GetHashCodeValue();
            int sig_hash = sighashK + sighashV;


            S_ESPARAMS es_params;
            if (!hashtable.ContainsKey(sig_hash))
                throw new Exception("Calibrate first");
            else
                es_params = (S_ESPARAMS)hashtable[sig_hash];

            signature = signature.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
            dict_rank map_signature = signature.MapRank();
            rank_dict signature_map = signature.RankMap();

            NormalDistribution norm = new NormalDistribution();

            string[] signature_genes = signature.Keys.ToArray();
            List<string> gsets = new List<string>();
            // from here enrichment analysis
            double[] sigvalues = signature.Values.ToArray();
            sigvalues = sigvalues.Plus(Pmult(norm.Generate(sigvalues.Length), 1 / (sigvalues.Average() * 10000)));
            double[] abs_signature = sigvalues.Abs();

            foreach (string key in library.Keys)
            {
                string[] gene_set = library[key];
                string[] stripped_set = strip_gene_set(signature_genes, gene_set);                
                if (stripped_set.Length >= min_size && stripped_set.Length <= max_size)
                {
                    int gsHash = stripped_set.GetHashCodeValue();
                    if (!hashgsea.ContainsKey(gsHash))                    
                        hashgsea[gsHash] = gsea_calc_es(abs_signature, map_signature, signature_map, es_params, stripped_set);                                            
                }
                string[] gene_set_pos = stripped_set.Where(k => dataset[k].FC > 0).ToArray();
                string[] stripped_set_pos = strip_gene_set(signature_genes, gene_set_pos);
                if (stripped_set_pos.Length >= min_size && stripped_set_pos.Length <= max_size)
                {
                    int gsHash = stripped_set_pos.GetHashCodeValue();
                    if (!hashgsea.ContainsKey(gsHash))
                        hashgsea[gsHash] = gsea_calc_es(abs_signature, map_signature, signature_map, es_params, stripped_set_pos);                  
                }
                string[] gene_set_neg = stripped_set.Where(k => dataset[k].FC < 0).ToArray();
                string[] stripped_set_neg = strip_gene_set(signature_genes, gene_set_neg);                
                if (stripped_set_neg.Length >= min_size && stripped_set_neg.Length <= max_size)
                {
                    int gsHash = stripped_set_neg.GetHashCodeValue();
                    if (!hashgsea.ContainsKey(gsHash))
                        hashgsea[gsHash] = gsea_calc_es(abs_signature, map_signature, signature_map, es_params, stripped_set_neg);
                                      
                }
            }


            // perform FDR etc.. later
            
        }


        public static S_GSEA gsea_calc(double[] abs_signature, string[] signature_genes, dict_rank map_signature, rank_dict signature_map, IEnumerable<string> geneset, S_ESPARAMS calibrated_model, ref Hashtable hashgsea, int min_size = 5, int max_size = 2000)
        {            

            string[] gene_set = geneset.ToArray();
            string[] stripped_set = strip_gene_set(signature_genes, gene_set);
            int gsHash = stripped_set.GetHashCodeValue();
            if (stripped_set.Length >= min_size && stripped_set.Length <= max_size)
            {                
                if (!hashgsea.ContainsKey(gsHash))
                    hashgsea[gsHash] = gsea_calc_es(abs_signature, map_signature, signature_map, calibrated_model, stripped_set);
            }

            return hashgsea.ContainsKey(gsHash) ? (S_GSEA)hashgsea[gsHash] : new S_GSEA() { pval = double.NaN };
        }


    }
}