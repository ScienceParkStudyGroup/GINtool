using System.Collections.Generic;

namespace GINtool
{
    internal class BsuRegulons
    {
        public const double NO_FC = -12345.6;
        public List<string> REGULONS = new List<string>();
        public string BSU = "";
        public double FC = NO_FC;
        public int NRUP = 0;
        public int NRDOWN = 0;
        public int NET = 0;
        public int TOT = 0;

        public BsuRegulons(double aFC, string aBSU)
        {
            BSU = aBSU;
            REGULONS = new List<string>();
            FC = aFC;
        }
        public BsuRegulons(string aBSU)
        {
            BSU = aBSU;
            REGULONS = new List<string>();
            FC = NO_FC;
        }
    }
}