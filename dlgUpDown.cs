using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

namespace GINtool
{
    public partial class dlgUpDown : Form
    {

        readonly List<string> pAvailItems = null;
        readonly List<string> pUpItems = null;
        readonly List<string> pDownItems = null;

        public dlgUpDown()
        {
            InitializeComponent();
        }


        public dlgUpDown(List<string> pAvail, List<string> pUp, List<string> pDown) : this()
        {
            pAvailItems = pAvail;
            pUpItems = pUp;
            pDownItems = pDown;

            if (pAvailItems != null)
                foreach (string it in pAvailItems)
                    lbAvail.Items.Add(it);

            if (pUpItems != null)
                foreach (string it in pUpItems)
                    lbUp.Items.Add(it);
            if (pDownItems != null)
                foreach (string it in pDownItems)
                    lbDown.Items.Add(it);

        }

        private void CopyFromTo(ListBox lbFrom, ListBox lbTo)
        {
            ArrayList moves = new ArrayList();

            foreach (object it in lbFrom.SelectedItems)
                moves.Add(it);

            for (int i = 0; i < moves.Count; i++)
            {
                lbFrom.Items.Remove(moves[i]);
                lbTo.Items.Add(moves[i]);

                if (lbFrom.Equals(lbAvail))
                    pAvailItems.Remove(moves[i].ToString());
                else
                    pAvailItems.Add(moves[i].ToString());

                if (lbFrom.Equals(lbUp))
                    pUpItems.Remove(moves[i].ToString());

                if (lbFrom.Equals(lbDown))
                    pDownItems.Remove(moves[i].ToString());


                if (lbTo.Equals(lbUp))
                    pUpItems.Add(moves[i].ToString());

                if (lbTo.Equals(lbDown))
                    pDownItems.Add(moves[i].ToString());
            }
        }

        private void btToUP_Click(object sender, EventArgs e)
        {
            CopyFromTo(lbAvail, lbUp);
        }

        private void btFromUp_Click(object sender, EventArgs e)
        {
            CopyFromTo(lbUp, lbAvail);
        }

        private void btToDown_Click(object sender, EventArgs e)
        {
            CopyFromTo(lbAvail, lbDown);
        }

        private void btFromDown_Click(object sender, EventArgs e)
        {
            CopyFromTo(lbDown, lbAvail);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Dispose();
            Close();
        }
    }
}
