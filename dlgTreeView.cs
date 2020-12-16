﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SysData = System.Data;

namespace GINtool
{
    public partial class dlgTreeView : Form
    {
        // define category columns 
        string[] catcols = new string[] { "cat1", "cat2", "cat3", "cat4", "cat5" };
        string[] regColumn = new string[] { Properties.Settings.Default.referenceRegulon };
        //string[] refColumn = null;
        bool catMode = true;

        List<cat_elements> gSelection = new List<cat_elements>();

        public List<cat_elements> GetSelection()
        {
            return gSelection;
        }

        public dlgTreeView()
        {
            InitializeComponent();
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

           
            if (catMode ? lvl < 5 : lvl<1)
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
                    if(__lcats.Rows.Count>0)
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
                TreeNode _node = (TreeNode) tn.Clone();
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
           
            if(treeNode != null)
            {                
                treeView2.Nodes.Add(treeNode);
                addToSelection(treeNode);
            }

        }

        // button unselect pressed
        private void button2_Click(object sender, EventArgs e)
        {
            TreeNode treeNode = treeView2.SelectedNode;
            if (treeNode.Parent != null)
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

                        TreeNode _tnode = headNode.Nodes[Int32.Parse(lvl[0])-1];
                        for (int i = 1; i < lvl.Count()-1; i++)
                            _tnode = _tnode.Nodes[Int32.Parse(lvl[i]) - 1];

                        _tnode.Nodes.Insert(Int32.Parse(lvl[lvl.Count()-1]) - 1,tnode);                                            
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
                    tnode.Nodes.Insert(index,node);
                    found = true;
                    break;
                }
                if(!found)
                    found = insertInChild(tnode, index,path, node);
            }
            return found;
        }


        public bool insertInChild(TreeNode original,int  index,string path, TreeNode node)
        {
            bool found = false;
            foreach (TreeNode tnode in original.Nodes)
            {
                if (tnode.FullPath == path)
                {
                    tnode.Nodes.Insert(index,node);
                    found = true;
                    break;
                }
                if(!found)
                    found = insertInChild(tnode, index,path, node);
            }
            return found;
        }

        // strip tag from positional info .. actually overlaps with checkParentNodes .. needs to be combined later
        private (TreeNode, string, int) getPositionInfo (TreeNode treeNode)
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
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
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
        public float[] fc;
        public float average;
    };

}