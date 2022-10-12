
namespace GINtool
{
    partial class dlgSelectData
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
            this.cbHeader = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbP = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tbFC = new System.Windows.Forms.TextBox();
            this.tbBSU = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btOk = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbHeader
            // 
            this.cbHeader.AutoSize = true;
            this.cbHeader.Checked = true;
            this.cbHeader.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbHeader.Location = new System.Drawing.Point(245, 25);
            this.cbHeader.Name = "cbHeader";
            this.cbHeader.Size = new System.Drawing.Size(124, 17);
            this.cbHeader.TabIndex = 3;
            this.cbHeader.Text = "my data has headers";
            this.cbHeader.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 113);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "p-value";
            // 
            // tbP
            // 
            this.tbP.Location = new System.Drawing.Point(105, 110);
            this.tbP.Name = "tbP";
            this.tbP.Size = new System.Drawing.Size(170, 20);
            this.tbP.TabIndex = 12;
            this.tbP.Validated += new System.EventHandler(this.tbP_Validated);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tbP);
            this.groupBox2.Controls.Add(this.tbFC);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.tbBSU);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(45, 59);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(304, 156);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "mandatory fields";
            // 
            // tbFC
            // 
            this.tbFC.Location = new System.Drawing.Point(105, 73);
            this.tbFC.Name = "tbFC";
            this.tbFC.Size = new System.Drawing.Size(170, 20);
            this.tbFC.TabIndex = 11;
            this.tbFC.Validated += new System.EventHandler(this.tbFC_Validated);
            // 
            // tbBSU
            // 
            this.tbBSU.Location = new System.Drawing.Point(105, 27);
            this.tbBSU.Name = "tbBSU";
            this.tbBSU.Size = new System.Drawing.Size(170, 20);
            this.tbBSU.TabIndex = 10;
            this.tbBSU.Validated += new System.EventHandler(this.tbBSU_Validated);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "fold-change";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "gene id";
            // 
            // btOk
            // 
            this.btOk.Location = new System.Drawing.Point(307, 254);
            this.btOk.Name = "btOk";
            this.btOk.Size = new System.Drawing.Size(75, 23);
            this.btOk.TabIndex = 9;
            this.btOk.Text = "OK";
            this.btOk.UseVisualStyleBackColor = true;
            this.btOk.Click += new System.EventHandler(this.btOk_Click);
            // 
            // btCancel
            // 
            this.btCancel.Location = new System.Drawing.Point(224, 254);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 10;
            this.btCancel.Text = "cancel";
            this.btCancel.UseVisualStyleBackColor = true;
            this.btCancel.Click += new System.EventHandler(this.btCancel_Click);
            // 
            // dlgSelectData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 301);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btOk);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.cbHeader);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "dlgSelectData";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "select data";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.CheckBox cbHeader;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btOk;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.TextBox tbBSU;
        private System.Windows.Forms.TextBox tbP;
        private System.Windows.Forms.TextBox tbFC;
    }
}