
namespace GINtool
{
    partial class dlgTreeView
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.treeView2 = new System.Windows.Forms.TreeView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.btnAllSel = new System.Windows.Forms.Button();
            this.btnAllBack = new System.Windows.Forms.Button();
            this.cbTableOutput = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.cbTopFC = new System.Windows.Forms.CheckBox();
            this.cbTopP = new System.Windows.Forms.CheckBox();
            this.udTOPP = new System.Windows.Forms.NumericUpDown();
            this.udTopFC = new System.Windows.Forms.NumericUpDown();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.cbCat = new System.Windows.Forms.CheckBox();
            this.udCat = new System.Windows.Forms.DomainUpDown();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.udTOPP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.udTopFC)).BeginInit();
            this.groupBox6.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.Location = new System.Drawing.Point(3, 16);
            this.treeView1.Name = "treeView1";
            this.treeView1.ShowNodeToolTips = true;
            this.treeView1.Size = new System.Drawing.Size(233, 397);
            this.treeView1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(309, 28);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "->";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(309, 57);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "<-";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(6, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(132, 20);
            this.textBox1.TabIndex = 4;
            // 
            // treeView2
            // 
            this.treeView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView2.Location = new System.Drawing.Point(3, 16);
            this.treeView2.Name = "treeView2";
            this.treeView2.Size = new System.Drawing.Size(233, 397);
            this.treeView2.TabIndex = 5;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.treeView1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(239, 416);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Available categories/regulons";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBox1);
            this.groupBox2.Location = new System.Drawing.Point(270, 369);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(144, 56);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "        # items selected";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.treeView2);
            this.groupBox3.Location = new System.Drawing.Point(443, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(239, 416);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Selection";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(713, 402);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 9;
            this.button3.Text = "Ok";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(713, 373);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 10;
            this.button4.Text = "Cancel";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // btnAllSel
            // 
            this.btnAllSel.Location = new System.Drawing.Point(309, 105);
            this.btnAllSel.Name = "btnAllSel";
            this.btnAllSel.Size = new System.Drawing.Size(75, 23);
            this.btnAllSel.TabIndex = 11;
            this.btnAllSel.Text = ">>";
            this.btnAllSel.UseVisualStyleBackColor = true;
            this.btnAllSel.Click += new System.EventHandler(this.btnAllSel_Click);
            // 
            // btnAllBack
            // 
            this.btnAllBack.Location = new System.Drawing.Point(309, 134);
            this.btnAllBack.Name = "btnAllBack";
            this.btnAllBack.Size = new System.Drawing.Size(75, 23);
            this.btnAllBack.TabIndex = 12;
            this.btnAllBack.Text = "<<";
            this.btnAllBack.UseVisualStyleBackColor = true;
            this.btnAllBack.Click += new System.EventHandler(this.btnAllBack_Click);
            // 
            // cbTableOutput
            // 
            this.cbTableOutput.AutoSize = true;
            this.cbTableOutput.Location = new System.Drawing.Point(709, 28);
            this.cbTableOutput.Name = "cbTableOutput";
            this.cbTableOutput.Size = new System.Drawing.Size(94, 17);
            this.cbTableOutput.TabIndex = 13;
            this.cbTableOutput.Text = "output to table";
            this.cbTableOutput.UseVisualStyleBackColor = true;
            this.cbTableOutput.CheckedChanged += new System.EventHandler(this.cbTableOutput_CheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.udTOPP);
            this.groupBox4.Controls.Add(this.cbTopP);
            this.groupBox4.Location = new System.Drawing.Point(270, 312);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(144, 51);
            this.groupBox4.TabIndex = 14;
            this.groupBox4.TabStop = false;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.udTopFC);
            this.groupBox5.Controls.Add(this.cbTopFC);
            this.groupBox5.Location = new System.Drawing.Point(270, 253);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(144, 51);
            this.groupBox5.TabIndex = 15;
            this.groupBox5.TabStop = false;
            // 
            // cbTopFC
            // 
            this.cbTopFC.AutoSize = true;
            this.cbTopFC.Location = new System.Drawing.Point(10, -1);
            this.cbTopFC.Name = "cbTopFC";
            this.cbTopFC.Size = new System.Drawing.Size(57, 17);
            this.cbTopFC.TabIndex = 0;
            this.cbTopFC.Text = "top FC";
            this.cbTopFC.UseVisualStyleBackColor = true;
            this.cbTopFC.CheckedChanged += new System.EventHandler(this.cbTopFC_CheckedChanged);
            // 
            // cbTopP
            // 
            this.cbTopP.AutoSize = true;
            this.cbTopP.Location = new System.Drawing.Point(10, -1);
            this.cbTopP.Name = "cbTopP";
            this.cbTopP.Size = new System.Drawing.Size(85, 17);
            this.cbTopP.TabIndex = 0;
            this.cbTopP.Text = "top P-values";
            this.cbTopP.UseVisualStyleBackColor = true;
            this.cbTopP.CheckedChanged += new System.EventHandler(this.cbTopP_CheckedChanged);
            // 
            // udTOPP
            // 
            this.udTOPP.Increment = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.udTOPP.Location = new System.Drawing.Point(6, 23);
            this.udTOPP.Maximum = new decimal(new int[] {
            25,
            0,
            0,
            0});
            this.udTOPP.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.udTOPP.Name = "udTOPP";
            this.udTOPP.Size = new System.Drawing.Size(132, 20);
            this.udTOPP.TabIndex = 1;
            this.udTOPP.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // udTopFC
            // 
            this.udTopFC.Increment = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.udTopFC.Location = new System.Drawing.Point(6, 20);
            this.udTopFC.Maximum = new decimal(new int[] {
            25,
            0,
            0,
            0});
            this.udTopFC.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.udTopFC.Name = "udTopFC";
            this.udTopFC.Size = new System.Drawing.Size(132, 20);
            this.udTopFC.TabIndex = 1;
            this.udTopFC.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.udCat);
            this.groupBox6.Controls.Add(this.cbCat);
            this.groupBox6.Location = new System.Drawing.Point(270, 199);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(144, 51);
            this.groupBox6.TabIndex = 16;
            this.groupBox6.TabStop = false;
            // 
            // cbCat
            // 
            this.cbCat.AutoSize = true;
            this.cbCat.Location = new System.Drawing.Point(10, -1);
            this.cbCat.Name = "cbCat";
            this.cbCat.Size = new System.Drawing.Size(112, 17);
            this.cbCat.TabIndex = 0;
            this.cbCat.Text = "category selection";
            this.cbCat.UseVisualStyleBackColor = true;
            this.cbCat.CheckedChanged += new System.EventHandler(this.cbCat_CheckedChanged);
            // 
            // udCat
            // 
            this.udCat.Items.Add("I");
            this.udCat.Items.Add("II");
            this.udCat.Items.Add("III");
            this.udCat.Items.Add("IV");
            this.udCat.Location = new System.Drawing.Point(6, 20);
            this.udCat.Name = "udCat";
            this.udCat.Size = new System.Drawing.Size(132, 20);
            this.udCat.TabIndex = 1;
            // 
            // dlgTreeView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(815, 450);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.cbTableOutput);
            this.Controls.Add(this.btnAllBack);
            this.Controls.Add(this.btnAllSel);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "dlgTreeView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Select categories/regulons";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.udTOPP)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.udTopFC)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TreeView treeView2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button btnAllSel;
        private System.Windows.Forms.Button btnAllBack;
        private System.Windows.Forms.CheckBox cbTableOutput;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.NumericUpDown udTOPP;
        private System.Windows.Forms.CheckBox cbTopP;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.NumericUpDown udTopFC;
        private System.Windows.Forms.CheckBox cbTopFC;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.DomainUpDown udCat;
        private System.Windows.Forms.CheckBox cbCat;
    }
}