using System.Collections.Generic;

namespace GINtool
{
    internal class BsuRegulons
    {
        public const double NO_FC = -12345.6;
        public List<string> REGULONS = new List<string>();
        public string BSU = "";
        public double FC = NO_FC;
        //public int NRUP = 0;
        //public int NRDOWN = 0;
        //public int NET = 0;
        //public int TOT = 0;
        //public int fpUP = 0;
        //public int fpDOWN = 0;
        public List<int> UP = new List<int>();
        public List<int> DOWN = new List<int>();
        public List<int> fpUP = new List<int>();
        public List<int> fpDOWN = new List<int>();


        public int NRDOWN { get { return DOWN.Count; } }
        public int NRUP { get { return UP.Count; } }
        public int NET { get { return UP.Count - DOWN.Count; } }
        public int TOT { get { return REGULONS.Count; } }

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