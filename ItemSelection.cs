using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace GINtool
{
    internal class ItemSelection
    {
        // define category columns 
        static string[] catColumns = new string[] { "cat1", "cat2", "cat3", "cat4", "cat5" };
        static string[] regColumn = new string[] { Properties.Settings.Default.referenceRegulon };
        static TreeNode gNodes = new TreeNode();

        private static string[] NodeTags(TreeNodeCollection treeNodes)
        {
            // Print the node.  
            //System.Diagnostics.Debug.WriteLine(treeNode.Text);
            //MessageBox.Show(treeNode.Text);
            // Print each node recursively.  

            List<string> _tags = new List<string>();

            foreach (TreeNode tn in treeNodes)
            {
                TreeNode _node = (TreeNode)tn.Clone();
                while (_node.Nodes.Count > 0)
                    _node = _node.Nodes[0];
                _tags.Add(_node.Tag.ToString());
            }

            return _tags.ToArray();
        }
        private static cat_elements createCategoryItem(TreeNode treeNode)
        {
            cat_elements sel = new cat_elements();
            List<string> codes = new List<string>();
            if (treeNode.Nodes.Count == 0)
            {
                string[] tags = treeNode.Tag.ToString().Split('_');
                codes.Add(tags[0]);
            }
            else
            {

                foreach (TreeNode _tnode in treeNode.Nodes)
                {
                    string[] _codes = NodeTags(_tnode.Nodes);
                    if (_codes.Length > 0)
                        codes.AddRange(_codes);
                    else // is end node
                        codes.Add(_tnode.Tag.ToString());

                }
            }
            sel.catName = treeNode.Text;
            sel.elTag = treeNode.Tag.ToString();
            sel.elements = codes.ToArray();

            return sel;
        }
        private static DataTable GetDistinctRecords(DataTable dt, string[] Columns)
        {
            return dt.DefaultView.ToTable(true, Columns);
        }
        public static TreeNode BuildMemTree(DataTable dt, bool catMode, TreeNode trv = null, int lvl = 1, string accumlevel = "")
        {
            // Clear the TreeView if there are another datas in this TreeView
            if (trv is null)
                trv = new TreeNode();

            DataTable _lcats = null;

            if (catMode)
                _lcats = GetDistinctRecords(dt, new string[] { catColumns[lvl - 1] });
            else
                _lcats = GetDistinctRecords(dt, regColumn);


            if (catMode ? lvl < 5 : lvl < 1)
            {
                int _rownr = 0;
                foreach (DataRow _row in _lcats.Rows)
                {

                    if (_row[0].ToString() == "")
                        return trv;

                    _rownr++;

                    TreeNode node = trv.Nodes.Add(_row[0].ToString());
                    // Set ToolTip text to reflect number of sub categories
                    //node.ToolTipText = string.Format("# subcat {0}", _lcats.Rows.Count.ToString());
                    // Store selection code in tag field
                    node.Tag = accumlevel == "" ? string.Format("{0}", _rownr) : string.Format("{0}.{1}", accumlevel, _rownr);

                    node.ToolTipText = node.Tag.ToString();

                    DataTable __lcats = dt.Select(string.Format("{0}='{1}'", catColumns[lvl - 1], node.Text)).CopyToDataTable();
                    if (__lcats.Rows.Count > 0)
                        BuildMemTree(__lcats, catMode, node, lvl: lvl + 1, accumlevel != "" ? accumlevel + "." + _rownr.ToString() : _rownr.ToString());
                }

            }
            else if (catMode) // level == 5 or 
            {
                foreach (DataRow _row in _lcats.Rows)
                {

                    if (_row[0].ToString() == "")
                        return null;

                    TreeNode _lNode = new TreeNode(_row[0].ToString());
                    _lNode.ToolTipText = string.Format("# subcat {0}", _lcats.Rows.Count.ToString());

                    return _lNode;
                }
            }
            else // !catMode
            {
                int _rownr = 0;
                foreach (DataRow _row in _lcats.Rows)
                {

                    if (_row[0].ToString() == "")
                        return null;

                    _rownr++;

                    TreeNode node = trv.Nodes.Add(_row[0].ToString());
                    node.Tag = accumlevel == "" ? string.Format("{0}", _rownr) : string.Format("{0}.{1}", accumlevel, _rownr);
                    node.ToolTipText = node.Tag.ToString();


                    //TreeNode _lNode = new TreeNode(_row[0].ToString());
                    //_lNode.ToolTipText = string.Format("# subcat {0}", _lcats.Rows.Count.ToString());

                    //return _lNode;
                }
            }

            return trv;
        }

        private static List<TreeNode> SelectCategoryLevel(int level)
        {
            List<TreeNode> selection = new List<TreeNode>();
            List<TreeNode> tmp_Selection1 = new List<TreeNode>();
            List<TreeNode> tmp_Selection2 = new List<TreeNode>();

            // start of with the default top level

            foreach (TreeNode node in gNodes.Nodes)
                tmp_Selection1.Add(node);
            selection = tmp_Selection1;


            // alternating tmp_selection1/tmp_selection2 for increasing levels.. to prevent nested loops

            if (level > 0)
            {
                foreach (TreeNode subnode in tmp_Selection1)
                {
                    foreach (TreeNode subsubnode in subnode.Nodes)
                        tmp_Selection2.Add(subsubnode);
                }
                selection = tmp_Selection2;
            }

            if (level > 1)
            {
                tmp_Selection1 = new List<TreeNode>();
                foreach (TreeNode subnode in tmp_Selection2)
                {
                    foreach (TreeNode subsubnode in subnode.Nodes)
                        tmp_Selection1.Add(subsubnode);
                }
                selection = tmp_Selection1;

            }
            if (level > 2)
            {
                tmp_Selection2 = new List<TreeNode>();
                foreach (TreeNode subnode in tmp_Selection1)
                {
                    foreach (TreeNode subsubnode in subnode.Nodes)
                        tmp_Selection2.Add(subsubnode);
                }
                selection = tmp_Selection2;
            }

            return selection;

        }

        public static List<TreeNode> SelectAllNodes(bool useCat = false)
        {
            List<TreeNode> selection = new List<TreeNode>();

            int maxCategories = catColumns.Length;

            if (!useCat)
            {
                foreach (TreeNode node in gNodes.Nodes)
                    selection.Add(node);
            }
            else
            {
                // coding level is reverse of order
                selection = SelectCategoryLevel(0);
            }

            return selection;

        }

        public static List<cat_elements> SelectAllElements(DataTable dataTable, bool catMode = true)
        {
            gNodes.Nodes.Clear();
            gNodes = BuildMemTree(dataTable, catMode, gNodes.Nodes.Add("Categories"), 1);

            List<TreeNode> selection = new List<TreeNode>();
            List<cat_elements> elements = new List<cat_elements>();

            if (!catMode)
            {
                foreach (TreeNode node in gNodes.Nodes)
                    selection.Add(node);
            }
            else
            {
                // coding level is reverse of order
                selection = SelectCategoryLevel(5);
            }

            foreach (TreeNode _t in selection)
                elements.Add(createCategoryItem(_t));

            return elements;
        }

        //public static void FillTree(DataTable dataTable, bool catMode = true)
        //{


        //    if (catMode)
        //    {
        //        catMode = true;
        //        lView = GetDistinctRecords(dataTable, catcols);
        //        //refColumn = catcols;
        //        BuildMemTree(dataTable, treeView1.Nodes.Add("Categories"), 1);
        //    }
        //    else
        //    {
        //        catMode = false;
        //        lView = GetDistinctRecords(dataTable, regColumn);
        //        //refColumn = regColumn;
        //        BuildMemTree(dataTable, treeView1.Nodes.Add("Regulons"), 1);
        //    }

        //}

    }
}
