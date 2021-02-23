using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SysData = System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace GINtool
{
    public partial class dlgTreeView : Form
    {
        // define category columns 
        string[] catcols = new string[] { "cat1", "cat2", "cat3", "cat4", "cat5" };
        string[] regColumn = new string[] { Properties.Settings.Default.referenceRegulon };
        //string[] refColumn = null;
        bool catMode = true;
        bool tableOutput = false;
        bool enableTableOutput = true;

        List<cat_elements> gSelection = new List<cat_elements>();

        public List<cat_elements> GetSelection()
        {
            return gSelection;
        }

        public dlgTreeView(bool categoryView = false, bool enableTable = true)
        {
            InitializeComponent();
            udCat.SelectedItem = udCat.Items[0];
            cbTopFC.Checked = false;
            cbTopP.Checked = false;
            udTopFC.Enabled = false;
            udTOPP.Enabled = false;
            cbCat.Checked = false;
            cbCat.Enabled = categoryView;
            udCat.Enabled = false;
            cbTableOutput.Enabled = enableTable;
        }

        // utility to select unique records
        private SysData.DataTable GetDistinctRecords(SysData.DataTable dt, string[] Columns)
        {
            return dt.DefaultView.ToTable(true, Columns);
        }

        // http://www.authorcode.com/create-treeview-from-datatable-in-c/


        //private TreeNode Searchnode(string nodetext, TreeView trv)
        //{
        //    foreach (TreeNode node in trv.Nodes)
        //    {
        //        if (node.Text == nodetext)
        //        {
        //            return node;
        //        }
        //    }

        //    return null;
        //}
        public void populateTree(SysData.DataTable dataTable, bool cat = true)
        {
            DataTable lView = null;
            if (cat)
            {
                catMode = true;
                lView = GetDistinctRecords(dataTable, catcols);
                //refColumn = catcols;
                BuildTree(dataTable, treeView1.Nodes.Add("Categories"), 1);
            }
            else
            {
                catMode = false;
                lView = GetDistinctRecords(dataTable, regColumn);
                //refColumn = regColumn;
                BuildTree(dataTable, treeView1.Nodes.Add("Regulons"), 1);
            }

        }

        public bool selectTableOutput()
        {
            return tableOutput;
        }


        public int getTopFC()
        {
            return cbTopFC.Checked ? (int)udTopFC.Value : -1;
        }

        public int getTopP()
        {
            return cbTopP.Checked ? (int)udTOPP.Value : -1;
        }



        private DataTable GetDistinctRegulons(DataTable dataTable, string[] regColumn)
        {
            throw new NotImplementedException();
        }



        // recursive population of treeview control
        public TreeNode BuildTree(DataTable dt, TreeNode trv = null, int lvl = 1, string accumlevel = "")
        {
            // Clear the TreeView if there are another datas in this TreeView
            if (trv is null)
                trv = new TreeNode();



            DataTable _lcats = null;

            if (catMode)
                _lcats = GetDistinctRecords(dt, new string[] { catcols[lvl - 1] });
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

                    DataTable __lcats = dt.Select(string.Format("{0}='{1}'", catcols[lvl - 1], node.Text)).CopyToDataTable();
                    if (__lcats.Rows.Count > 0)
                        BuildTree(__lcats, node, lvl: lvl + 1, accumlevel != "" ? accumlevel + "." + _rownr.ToString() : _rownr.ToString());

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

        private string[] nodesToArray(TreeNode treeNode)
        {
            return null;
        }

        private string[] NodeTags(TreeNodeCollection treeNodes)
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

        private cat_elements createCategoryItem(TreeNode treeNode)
        {
            cat_elements sel;
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

        private void addToSelection(TreeNode treeNode)
        {
            gSelection.Add(createCategoryItem(treeNode));
        }

        private void removeFromSelection(int index)
        {
            gSelection.RemoveAt(index);
        }


        // select button pressed
        private void button1_Click(object sender, EventArgs e)
        {
            if(treeView1.SelectedNode is null)
                return;

            if (treeView1.SelectedNode.Parent == null)
            {
                MessageBox.Show("Cannot select top node");
                return;
            }

            TreeNode treeNode = treeView1.SelectedNode;
            if (treeNode == null)
                return;

            string fp = treeNode.FullPath;
            int nodeIndex = treeNode.Index;
            // remove last part of tree
            int rstr = fp.LastIndexOf('\\');
            if (rstr > 0) fp = fp.Remove(rstr);
            // add position information in Tag field
            treeNode.Tag = treeNode.Tag + "_" + fp + "_" + nodeIndex.ToString();

            treeView1.Nodes.Remove(treeNode);

            if (treeNode != null)
            {
                treeView2.Nodes.Add(treeNode);
                addToSelection(treeNode);
            }
            UpdateCounter();
        }

        // button unselect pressed
        private void button2_Click(object sender, EventArgs e)
        {
            TreeNode treeNode = treeView2.SelectedNode;
            if (treeNode==null  || treeNode.Parent != null)
            {
                MessageBox.Show("Only main-nodes are allowed");
                return;
            }
            treeView2.Nodes.Remove(treeNode);
            (TreeNode node, string fullpath, int idx) = getPositionInfo(treeNode);
            if (node != null)
            {
                removeFromSelection(treeNode.Index);

                if (!insertNode(node, idx, fullpath)) // add as main node
                    treeView1.Nodes[0].Nodes.Add(node);
                else // check if main nodes are correctly placed
                    checkParentNodes();
            }

            UpdateCounter();
        }

        // check integrity of tree if new (sub) head node is added
        public void checkParentNodes()
        {
            TreeNode headNode = treeView1.Nodes[0];
            foreach (TreeNode tnode in headNode.Nodes)
            {
                //tnode.Nodes.Insert(index, node);
                if (tnode.Tag != null)
                {
                    string[] lvl = tnode.Tag.ToString().Split('.');
                    if (lvl.Count() > 1)
                    {
                        treeView1.Nodes.Remove(tnode);

                        TreeNode _tnode = headNode.Nodes[Int32.Parse(lvl[0]) - 1];
                        for (int i = 1; i < lvl.Count() - 1; i++) // need to fix this here... should check that nodes do not contain genes then this check is not necessary.
                            try
                            {
                                _tnode = _tnode.Nodes[Int32.Parse(lvl[i]) - 1];
                            }
                            catch
                            {

                            }

                        _tnode.Nodes.Insert(Int32.Parse(lvl[lvl.Count() - 1]) - 1, tnode);
                    }
                }

            }

        }

        // copied from inet, insert item in tree at specified location

        public bool insertInParent(string path, int index, TreeNode node)
        {
            bool found = false;
            foreach (TreeNode tnode in treeView1.Nodes)
            {
                if (tnode.FullPath == path)
                {
                    tnode.Nodes.Insert(index, node);
                    found = true;
                    break;
                }
                if (!found)
                    found = insertInChild(tnode, index, path, node);
            }
            return found;
        }


        public bool insertInChild(TreeNode original, int index, string path, TreeNode node)
        {
            bool found = false;
            foreach (TreeNode tnode in original.Nodes)
            {
                if (tnode.FullPath == path)
                {
                    tnode.Nodes.Insert(index, node);
                    found = true;
                    break;
                }
                if (!found)
                    found = insertInChild(tnode, index, path, node);
            }
            return found;
        }

        // strip tag from positional info .. actually overlaps with checkParentNodes .. needs to be combined later
        private (TreeNode, string, int) getPositionInfo(TreeNode treeNode)
        {
            string[] tags = treeNode.Tag.ToString().Split('_');
            treeNode.Tag = tags[0];
            return (treeNode, tags[1], Int32.Parse(tags[2]));
        }

        // insert in tree from treeview 1
        private bool insertNode(TreeNode treeNode, int index, string fullPath)
        {
            return insertInParent(fullPath, index, treeNode);

        }


        private void button3_Click(object sender, EventArgs e)
        {
            if (gSelection.Count() > 255 && (!cbTopFC.Checked && !cbTopP.Checked))
            {
                DialogResult dialogResult = MessageBox.Show("The number of series is larger than 255 and cannot be plotted, the output will be redirected to a table. Continue?", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    this.cbTableOutput.Checked = true;
                    //do something
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }
            }
            else
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }


        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            List<TreeNode> selection = new List<TreeNode>();

            foreach (TreeNode node in treeNode.Nodes)
                selection.Add(node);


            foreach (TreeNode node in selection)
            {
                string fp = node.FullPath;
                int nodeIndex = node.Index;
                // remove last part of tree
                int rstr = fp.LastIndexOf('\\');
                if (rstr > 0) fp = fp.Remove(rstr);
                // add position information in Tag field
                node.Tag = node.Tag + "_" + fp + "_" + nodeIndex.ToString();

                treeView1.Nodes.Remove(node);

                if (node != null)
                {
                    treeView2.Nodes.Add(node);
                    addToSelection(node);
                }

            }

        }

        // NOTE   This code can be added to the BeforeCheck event handler instead of the AfterCheck event.
        // After a tree node's Checked property is changed, all its child nodes are updated to the same value.
        private void node_AfterCheck(object sender, TreeViewEventArgs e)
        {
            // The code only executes if the user caused the checked state to change.
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    /* Calls the CheckAllChildNodes method, passing in the current 
                    Checked value of the TreeNode whose checked state changed. */
                    this.CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }


        private List<TreeNode> SelectCategoryLevel(int level)
        {
            List<TreeNode> selection = new List<TreeNode>();
            List<TreeNode> tmp_Selection1 = new List<TreeNode>();
            List<TreeNode> tmp_Selection2 = new List<TreeNode>();

            // start of with the default top level
            
            foreach (TreeNode node in treeView1.TopNode.Nodes)
                tmp_Selection1.Add(node);
            selection = tmp_Selection1;
            
            
            // alternating tmp_selection1/tmp_selection2 for increasing levels.. to prevent nested loops

            if (level>0)
            {
                foreach (TreeNode subnode in tmp_Selection1)
                {
                    foreach(TreeNode subsubnode in subnode.Nodes)
                        tmp_Selection2.Add(subsubnode);
                }
                selection = tmp_Selection2;
            }

            if (level >1)
            {
                tmp_Selection1 = new List<TreeNode>();
                foreach (TreeNode subnode in tmp_Selection2)
                {
                    foreach (TreeNode subsubnode in subnode.Nodes)
                        tmp_Selection1.Add(subsubnode);
                }
                selection = tmp_Selection1;

            }
            if(level>2)
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


        private void btnAllSel_Click(object sender, EventArgs e)
        {
            List<TreeNode> selection = new List<TreeNode>();

            if (!cbCat.Checked)

            {
                foreach (TreeNode node in treeView1.TopNode.Nodes)
                    selection.Add(node);
            }
            else
            {
                selection = SelectCategoryLevel(udCat.SelectedIndex);
            }            

            foreach (TreeNode node in selection)
            {
                string fp = node.FullPath;
                int nodeIndex = node.Index;
                // remove last part of tree
                int rstr = fp.LastIndexOf('\\');
                if (rstr > 0) fp = fp.Remove(rstr);
                // add position information in Tag field
                node.Tag = node.Tag + "_" + fp + "_" + nodeIndex.ToString();

                treeView1.Nodes.Remove(node);

                if (node != null)
                {
                    treeView2.Nodes.Add(node);
                    addToSelection(node);
                }

            }
            UpdateCounter();
        }

        void UpdateCounter()
        {
            textBox1.Text = treeView2.Nodes.Count.ToString();
        }

        private void btnAllBack_Click(object sender, EventArgs e)
        {
            List<TreeNode> selection = new List<TreeNode>();

            foreach (TreeNode node in treeView2.Nodes)
                selection.Add(node);

            foreach (TreeNode treeNode in selection)
            {
                treeView2.Nodes.Remove(treeNode);
                (TreeNode node, string fullpath, int idx) = getPositionInfo(treeNode);
                removeFromSelection(treeNode.Index);
                if (!insertNode(node, idx, fullpath)) // add as main node
                    treeView1.Nodes[0].Nodes.Add(node);
                else // check if main nodes are correctly placed
                    checkParentNodes();
            }

            UpdateCounter();
        }

        private void cbTableOutput_CheckedChanged(object sender, EventArgs e)
        {
            tableOutput = cbTableOutput.Checked;
        }

        private void cbCat_CheckedChanged(object sender, EventArgs e)
        {
            udCat.Enabled = cbCat.Checked;
        }

        private void cbTopFC_CheckedChanged(object sender, EventArgs e)
        {
            udTopFC.Enabled = cbTopFC.Checked;
            if (cbTopFC.Checked)
            {
                udTOPP.Enabled = false;
                cbTopP.Checked = false;
            }
        }

        private void cbTopP_CheckedChanged(object sender, EventArgs e)
        {
            udTOPP.Enabled = cbTopP.Checked;
            if (cbTopP.Checked)
            {
                udTopFC.Enabled = false;
                cbTopFC.Checked = false;
            }
        }
    }


    public struct cat_elements
    {
        public string catName;
        public string elTag;
        public string[] elements;
    };


    public struct element_fc
    {
        public string catName;        
        public double[] fc;
        public double[] pvalues;
        public double averagep;
        public double madp;
        public double average;
        public double sd;
        public double mad;
        public string[] genes;
    };

    public struct element_rank
    {
        public double[] average_fc;
        public double[] mad_fc;
        public string catName;
        public int[] nr_genes;
        public string[] genes;
    }


#if CLICK_CHART

    public struct chart_info
    {
    
        public chart_info(Excel.Chart cht, List<element_fc> els)
        {
            chart = cht;
            chartData = els;
        }
        public Excel.Chart chart;
        public List<element_fc> chartData;
    }
#endif

}
