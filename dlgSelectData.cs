using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GINtool
{
    public partial class dlgSelectData : Form
    {

        Excel.Range rangeBSU;
        Excel.Range rangeFC;
        Excel.Range rangeP;
        
        
        public Excel.Application theApp = null;
        protected override void OnLoad(EventArgs e)
        {
            var btn = new Button();
            btn.Size = new Size(25, tbBSU.ClientSize.Height + 2);
            btn.Location = new Point(tbBSU.ClientSize.Width - btn.Width, -1);
            btn.Cursor = Cursors.Default;
            btn.Image = Properties.Resources.keyboard;
            btn.Click += button1_Click;
            tbBSU.Controls.Add(btn);
            // Send EM_SETMARGINS to prevent text from disappearing underneath the button
            var btn2 = new Button();
            btn2.Size = new Size(25, tbFC.ClientSize.Height + 2);
            btn2.Location = new Point(tbFC.ClientSize.Width - btn.Width, -1);
            btn2.Cursor = Cursors.Default;
            btn2.Image = Properties.Resources.keyboard;
            btn2.Click += button2_Click;
            tbFC.Controls.Add(btn2);

            var btn3 = new Button();
            btn3.Size = new Size(25, tbFC.ClientSize.Height + 2);
            btn3.Location = new Point(tbFC.ClientSize.Width - btn.Width, -1);
            btn3.Cursor = Cursors.Default;
            btn3.Image = Properties.Resources.keyboard;
            btn3.Click += button3_Click;
            tbP.Controls.Add(btn3);

            SendMessage(tbBSU.Handle, 0xd3, (IntPtr)2, (IntPtr)(btn.Width << 16));
            base.OnLoad(e);
        }
      

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);


        private string RangeAddress(Excel.Range rng)
        {
            return rng.get_AddressLocal(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }
        private string CellAddress(Excel.Worksheet sht, int row, int col)
        {
            return RangeAddress(sht.Cells[row, col]);
        }

        private (string, int) cell2colint(string cell)
        {
            string[] _r = cell.Split('$');

            return (_r[1], Int32.Parse(_r[2]));
        }

        private string colint2cell(string col,int r)
        {
            return string.Format("${0}${1}", col, r.ToString());
        }
        private string stripheader(Excel.Range range)
        {
            string rstr = range.Address.ToString();
            (string firstc, int firstr) = cell2colint(rstr.Split(':')[0]);
            (string lastc, int lastr) = cell2colint(rstr.Split(':')[1]);

            string r = string.Format("{0}:{1}", colint2cell(firstc, firstr + 1), colint2cell(lastc, lastr));

            return r;
        }


        bool Checkoutput()
        {            
            
            if(rangeBSU.Rows.Count!=rangeFC.Rows.Count || rangeBSU.Rows.Count != rangeP.Rows.Count || rangeFC.Rows.Count != rangeP.Rows.Count)
            {
                MessageBox.Show("The input sizes of the selected columns is not the same. Please adjust!");
                return false;
            }
                
            if(rangeBSU.Address.ToString() == rangeFC.Address.ToString() || rangeBSU.Address.ToString() == rangeP.Address.ToString() || rangeFC.Address.ToString() == rangeP.Address.ToString())
            {
                MessageBox.Show("Some of the inputs are pointing to the same data. Please adjust!");
                return false;
            }

            return true;

        }


        static IEnumerable<string> GetExcelStrings()
        {
            string[] alphabet = { string.Empty, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            return from c1 in alphabet
                   from c2 in alphabet
                   from c3 in alphabet.Skip(1)                    // c3 is never empty
                   where c1 == string.Empty || c2 != string.Empty // only allow c2 to be empty if c1 is also empty
                   select c1 + c2 + c3;
        }

        public dlgSelectData(Excel.Range selection=null)
        {
            InitializeComponent();
            if(selection!=null && selection.Columns.Count==3)
            {
                string s1 = RangeAddress(selection);


                int startCol = selection.Column;
                int startRow = selection.Row;
                
                int nrRows = selection.Rows.Count;                


                tbBSU.Text = string.Format("{3}!${0}${1}:${0}${2}", GetExcelStrings().ElementAt(startCol + 1), startRow, startRow + nrRows-1,selection.Worksheet.Name);
                tbFC.Text = string.Format("{3}!${0}${1}:${0}${2}", GetExcelStrings().ElementAt(startCol), startRow, startRow + nrRows - 1, selection.Worksheet.Name);
                tbP.Text = string.Format("{3}!${0}${1}:${0}${2}", GetExcelStrings().ElementAt(startCol-1), startRow, startRow + nrRows - 1, selection.Worksheet.Name);


                rangeP = selection.Columns[1];
                rangeFC = selection.Columns[2];
                rangeBSU = selection.Columns[3];                                
            }
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btOk_Click(object sender, EventArgs e)
        {
            if (Checkoutput())
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }            
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            this.BringToFront();
            var range = theApp.InputBox(Prompt: "Select BSU codes", Title: "Select Range", Type: 8);
            this.Visible = true;

            if (range.GetType().ToString() == "System.Boolean")
            {
                return ;
            }

            if(range == null || ((Excel.Range)range).Columns.Count>1)
            {
                MessageBox.Show("Please select a single column");
                return;
            }

            rangeBSU = (Excel.Range)range;

            if (cbHeader.Checked)
                tbBSU.Text = string.Format("{0}!{1}", rangeBSU.Worksheet.Name, stripheader(rangeBSU));
            else
                tbBSU.Text = string.Format("{0}!{1}", rangeBSU.Worksheet.Name, rangeBSU.Address.ToString());
            
        }

        private Excel.Range GetAdjustedColumn(Excel.Range colData)
        {
            string rstr = colData.Address.ToString();
            (string firstc, int firstr) = cell2colint(rstr.Split(':')[0]);
            (string lastc, int lastr) = cell2colint(rstr.Split(':')[1]);

            int offset = 0;
            if (cbHeader.Checked)
                offset = 1;

            string stStart = colint2cell(firstc, firstr+offset);
            string stdEnd = colint2cell(lastc, lastr);

          
            Excel.Range startCell = colData.Worksheet.Range[stStart];
            Excel.Range endCell = colData.Worksheet.Range[stdEnd];

            return colData.Worksheet.Range[startCell, endCell];


        }


        public Excel.Range getBSU()
        {
            return GetAdjustedColumn(rangeBSU);
        }

        public Excel.Range getP()
        {
            return GetAdjustedColumn(rangeP);            
        }

        public Excel.Range getFC()
        {
            return GetAdjustedColumn(rangeFC);         
        }




        private void button3_Click(object sender, EventArgs e)
        {

            this.Visible = false;
            this.BringToFront();
            var range = theApp.InputBox(Prompt: "Select p-values", Title: "Select Range",  Type: 8);
            this.Visible = true;

            if (range.GetType().ToString() == "System.Boolean")
            {
                return;
            }

            if (range == null || range.Columns.Count > 1)
            {
                MessageBox.Show("Please select a single column");
                return;
            }

            rangeP = range;

            if (cbHeader.Checked)
                tbP.Text = string.Format("{0}!{1}", range.Worksheet.Name, stripheader(rangeP)); 
            else
                tbP.Text = string.Format("{0}!{1}", range.Worksheet.Name, rangeP.Address.ToString()); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            this.BringToFront();
            var range = theApp.InputBox(Prompt: "Select fold changes", Title: "Select Range", Type: 8);
            this.Visible = true;
            if (range.GetType().ToString() == "System.Boolean")
            {
                return;
            }
            if (range == null || range.Columns.Count > 1)
            {
                MessageBox.Show("Please select a single column");
                return;
            }

            rangeFC = range;

            if (cbHeader.Checked)
                tbFC.Text = string.Format("{0}!{1}", range.Worksheet.Name, stripheader(rangeFC)); 
            else
                tbFC.Text = string.Format("{0}!{1}", range.Worksheet.Name, rangeFC.Address.ToString()); 
        }


        private (bool, Excel.Range) ValidateTextBox2(TextBox tb, Excel.Range range)
        {
            bool result = false;
            Excel.Range lRange = range;
            
            string rstr = tb.Text;
            (string firstc, int firstr) = cell2colint(rstr.Split(':')[0]);
            (string lastc, int lastr) = cell2colint(rstr.Split(':')[1]);


            if (firstc != lastc)
            {
                MessageBox.Show("Please select only a single column");
                return (result, lRange);
            }

            int offset = 0;
            if (cbHeader.Checked)
                offset = 1;

            string stStart = colint2cell(firstc, firstr + offset);
            string stdEnd = colint2cell(lastc, lastr);

            try
            {
                if (lRange != null)
                {
                    Excel.Range startCell;
                    Excel.Range endCell;
                    try
                    {
                        startCell = lRange.Worksheet.Range[stStart];
                        endCell = lRange.Worksheet.Range[stdEnd];
                        Excel.Range tmpRange = lRange.Worksheet.Range[startCell, endCell];

                        lRange = tmpRange;
                        result = true;
                    }
                    catch 
                    {
                        tb.Text = string.Format("{0}!{1)", lRange.Worksheet.Name, lRange.Address.ToString());
                        MessageBox.Show("You entered an invalid range");
                    }

                }
            }
            catch 
            {
                tb.Text = string.Format("{0}!{1)", lRange.Worksheet.Name, lRange.Address.ToString());
                MessageBox.Show("You entered an invalid range");

            }

            return (result, lRange);

        }

     

        private void tbBSU_Validated(object sender, EventArgs e)
        {
            (bool result, Excel.Range lRange) = ValidateTextBox2(tbBSU, rangeBSU);
            if (result)
            {
                rangeBSU = lRange;
                tbBSU.Text = string.Format("{0}!{1}", rangeBSU.Worksheet.Name, rangeBSU.Address.ToString());
            }
            else
            {
                if(rangeBSU!=null)
                    tbBSU.Text = string.Format("{0}!{1}", rangeBSU.Worksheet.Name, rangeBSU.Address.ToString());
            }
        }

        private void tbFC_Validated(object sender, EventArgs e)
        {
            (bool result, Excel.Range lRange) = ValidateTextBox2(tbFC, rangeFC);
            if (result)
            {
                rangeFC = lRange;
                tbFC.Text = string.Format("{0}!{1}", rangeFC.Worksheet.Name, rangeFC.Address.ToString());
            }
            else
            {
                if (rangeFC != null)
                    tbFC.Text = string.Format("{0}!{1}", rangeFC.Worksheet.Name, rangeFC.Address.ToString());
            }
        }

        private void tbP_Validated(object sender, EventArgs e)
        {
            (bool result, Excel.Range lRange) = ValidateTextBox2(tbP, rangeP);
            if (result)
            {
                rangeP = lRange;
                tbP.Text = string.Format("{0}!{1}", rangeP.Worksheet.Name, rangeP.Address.ToString());
            }
            else
            {
                if (rangeP != null)
                    tbP.Text = string.Format("{0}!{1}", rangeP.Worksheet.Name, rangeP.Address.ToString());
            }
        }
    }



}
