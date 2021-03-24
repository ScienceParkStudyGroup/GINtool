using System.Collections.Generic;

namespace GINtool
{
    internal class BsuRegulons
    {

        // a gene can be regulated by more then 1 regulon..furthermore different regulons can influence the gene differently (i.e. up and down regulation)

        public const double NO_FC = -12345.6;
        public const double NO_PVALUE = -654321.1;
        public double PVALUE = NO_PVALUE;
        public List<string> REGULONS = new List<string>();
        public string BSU = "";
        public double FC = NO_FC;

        public string GENE = "";

        public List<int> UP = new List<int>();
        public List<int> DOWN = new List<int>();
        

        public int NRDOWN { get { return DOWN.Count; } }
        public int NRUP { get { return UP.Count; } }
        public int NET { get { return UP.Count - DOWN.Count; } }
        public int TOT { get { return REGULONS.Count; } }

        public BsuRegulons(double aFC, double aPvalue, string aBSU)
        {
            BSU = aBSU;
            PVALUE = aPvalue;
            REGULONS = new List<string>();
            FC = aFC;
            GENE = "";
        }
        public BsuRegulons(string aBSU)
        {
            BSU = aBSU;
            REGULONS = new List<string>();
            FC = NO_FC;
            PVALUE = NO_PVALUE;
            GENE = "";
        }

        public BsuRegulons()
        {
            BSU = "";
            REGULONS = new List<string>();
            FC = NO_FC;
            PVALUE = NO_PVALUE;
            GENE = "";
        }
    }
}