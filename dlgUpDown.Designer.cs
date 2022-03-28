namespace GINtool
{
    partial class dlgUpDown
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
            this.gpAvail = new System.Windows.Forms.GroupBox();
            this.lbAvail = new System.Windows.Forms.ListBox();
            this.gpUP = new System.Windows.Forms.GroupBox();
            this.lbUp = new System.Windows.Forms.ListBox();
            this.gbDown = new System.Windows.Forms.GroupBox();
            this.lbDown = new System.Windows.Forms.ListBox();
            this.btToUP = new System.Windows.Forms.Button();
            this.btFromUp = new System.Windows.Forms.Button();
            this.btToDown = new System.Windows.Forms.Button();
            this.btFromDown = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.gpAvail.SuspendLayout();
            this.gpUP.SuspendLayout();
            this.gbDown.SuspendLayout();
            this.SuspendLayout();
            // 
            // gpAvail
            // 
            this.gpAvail.Controls.Add(this.lbAvail);
            this.gpAvail.Location = new System.Drawing.Point(12, 12);
            this.gpAvail.Name = "gpAvail";
            this.gpAvail.Size = new System.Drawing.Size(200, 349);
            this.gpAvail.TabIndex = 1;
            this.gpAvail.TabStop = false;
            this.gpAvail.Text = "undefined";
            // 
            // lbAvail
            // 
            this.lbAvail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbAvail.FormattingEnabled = true;
            this.lbAvail.Location = new System.Drawing.Point(3, 16);
            this.lbAvail.Name = "lbAvail";
            this.lbAvail.Size = new System.Drawing.Size(194, 330);
            this.lbAvail.TabIndex = 0;
            // 
            // gpUP
            // 
            this.gpUP.Controls.Add(this.lbUp);
            this.gpUP.Location = new System.Drawing.Point(324, 12);
            this.gpUP.Name = "gpUP";
            this.gpUP.Size = new System.Drawing.Size(200, 162);
            this.gpUP.TabIndex = 2;
            this.gpUP.TabStop = false;
            this.gpUP.Text = "up-regulated";
            // 
            // lbUp
            // 
            this.lbUp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbUp.FormattingEnabled = true;
            this.lbUp.Location = new System.Drawing.Point(3, 16);
            this.lbUp.Name = "lbUp";
            this.lbUp.Size = new System.Drawing.Size(194, 143);
            this.lbUp.TabIndex = 0;
            // 
            // gbDown
            // 
            this.gbDown.Controls.Add(this.lbDown);
            this.gbDown.Location = new System.Drawing.Point(324, 196);
            this.gbDown.Name = "gbDown";
            this.gbDown.Size = new System.Drawing.Size(200, 162);
            this.gbDown.TabIndex = 3;
            this.gbDown.TabStop = false;
            this.gbDown.Text = "down-regulated";
            // 
            // lbDown
            // 
            this.lbDown.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbDown.FormattingEnabled = true;
            this.lbDown.Location = new System.Drawing.Point(3, 16);
            this.lbDown.Name = "lbDown";
            this.lbDown.Size = new System.Drawing.Size(194, 143);
            this.lbDown.TabIndex = 0;
            // 
            // btToUP
            // 
            this.btToUP.Location = new System.Drawing.Point(226, 64);
            this.btToUP.Name = "btToUP";
            this.btToUP.Size = new System.Drawing.Size(75, 23);
            this.btToUP.TabIndex = 4;
            this.btToUP.Text = "->";
            this.btToUP.UseVisualStyleBackColor = true;
            this.btToUP.Click += new System.EventHandler(this.btToUP_Click);
            // 
            // btFromUp
            // 
            this.btFromUp.Location = new System.Drawing.Point(226, 93);
            this.btFromUp.Name = "btFromUp";
            this.btFromUp.Size = new System.Drawing.Size(75, 23);
            this.btFromUp.TabIndex = 5;
            this.btFromUp.Text = "<-";
            this.btFromUp.UseVisualStyleBackColor = true;
            this.btFromUp.Click += new System.EventHandler(this.btFromUp_Click);
            // 
            // btToDown
            // 
            this.btToDown.Location = new System.Drawing.Point(226, 247);
            this.btToDown.Name = "btToDown";
            this.btToDown.Size = new System.Drawing.Size(75, 23);
            this.btToDown.TabIndex = 6;
            this.btToDown.Text = "->";
            this.btToDown.UseVisualStyleBackColor = true;
            this.btToDown.Click += new System.EventHandler(this.btToDown_Click);
            // 
            // btFromDown
            // 
            this.btFromDown.Location = new System.Drawing.Point(226, 276);
            this.btFromDown.Name = "btFromDown";
            this.btFromDown.Size = new System.Drawing.Size(75, 23);
            this.btFromDown.TabIndex = 7;
            this.btFromDown.Text = "<-";
            this.btFromDown.UseVisualStyleBackColor = true;
            this.btFromDown.Click += new System.EventHandler(this.btFromDown_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(449, 378);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "Close";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dlgUpDown
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(537, 426);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btFromDown);
            this.Controls.Add(this.btToDown);
            this.Controls.Add(this.btFromUp);
            this.Controls.Add(this.btToUP);
            this.Controls.Add(this.gbDown);
            this.Controls.Add(this.gpUP);
            this.Controls.Add(this.gpAvail);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "dlgUpDown";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "up/down regulation mapping";
            this.gpAvail.ResumeLayout(false);
            this.gpUP.ResumeLayout(false);
            this.gbDown.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox gpAvail;
        private System.Windows.Forms.GroupBox gpUP;
        private System.Windows.Forms.GroupBox gbDown;
        private System.Windows.Forms.Button btToUP;
        private System.Windows.Forms.Button btFromUp;
        private System.Windows.Forms.Button btToDown;
        private System.Windows.Forms.Button btFromDown;
        private System.Windows.Forms.ListBox lbAvail;
        private System.Windows.Forms.ListBox lbUp;
        private System.Windows.Forms.ListBox lbDown;
        private System.Windows.Forms.Button button1;
    }
}