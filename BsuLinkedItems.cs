using System.Collections.Generic;

namespace GINtool
{

    internal class CategoryItem
    {
        public string Name;
        public string Id;
        public string catID;
        public CategoryItem(string aName, string aCatID, string aId) { Name = aName; catID = aCatID; Id = aId; }

    };

    internal class DataItem
    {
        public double FC, pval;
        public string BSU;
    }

    internal class RegulonItem
    {
        public string Name;
        public string Direction;
        public RegulonItem(string aName, string aDirection) { Name = aName; Direction = aDirection; }
    }

    internal class BsuLinkedItems
    {

        // a gene can be regulated by more then 1 regulon..furthermore different regulons can influence the gene differently (i.e. up and down regulation)

        public const double NO_FC = -12345.6;
        public const double NO_PVALUE = -654321.1;
        public double PVALUE = NO_PVALUE;

        public List<RegulonItem> Regulons = new List<RegulonItem>();
        public List<CategoryItem> Categories = new List<CategoryItem>();
        public string BSU = "";
        public double FC = NO_FC;

        public List<int> REGULON_UP = new List<int>();
        public List<int> REGULON_DOWN = new List<int>();
        public List<int> REGULON_UNKNOWN_DIR = new List<int>();

        public string GeneName = "";
        public string GeneDescription = "";
        public string GeneFunction = "";

        public int REGULON_NRDOWN { get { return REGULON_DOWN.Count; } }
        public int REGULON_NRUP { get { return REGULON_UP.Count; } }
        public int REGULON_NET { get { return REGULON_UP.Count - REGULON_DOWN.Count; } }
        public int REGULON_TOT { get { return Regulons.Count; } }

        public BsuLinkedItems(double aFC, double aPvalue, string aBSU)
        {
            BSU = aBSU;
            PVALUE = aPvalue;
            Regulons = new List<RegulonItem>();
            Categories = new List<CategoryItem>();
            FC = aFC;
            GeneName = "";
        }
        public BsuLinkedItems(string aBSU)
        {
            BSU = aBSU;
            Regulons = new List<RegulonItem>();
            Categories = new List<CategoryItem>();
            FC = NO_FC;
            PVALUE = NO_PVALUE;
            GeneName = "";
            GeneFunction = "";
            GeneDescription = "";
        }

        public BsuLinkedItems()
        {
            BSU = "";
            Regulons = new List<RegulonItem>();
            Categories = new List<CategoryItem>();
            FC = NO_FC;
            PVALUE = NO_PVALUE;
            GeneName = "";
            GeneFunction = "";
            GeneDescription = "";
        }
    }
}