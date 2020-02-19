namespace SR1000
{
    partial class Paging
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label2 = new System.Windows.Forms.Label();
            this.txtPageSize = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblCurrentPage = new System.Windows.Forms.Label();
            this.lblSept = new System.Windows.Forms.Label();
            this.lblPageCount = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblTotalCount = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.linkFirst = new System.Windows.Forms.Button();
            this.linkPrevious = new System.Windows.Forms.Button();
            this.linkNext = new System.Windows.Forms.Button();
            this.linkLast = new System.Windows.Forms.Button();
            this.txtPageNum = new System.Windows.Forms.TextBox();
            this.btnGo = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(15, 11);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 25);
            this.label2.TabIndex = 246;
            this.label2.Text = "每页";
            // 
            // txtPageSize
            // 
            this.txtPageSize.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.txtPageSize.Location = new System.Drawing.Point(75, 11);
            this.txtPageSize.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPageSize.Name = "txtPageSize";
            this.txtPageSize.Size = new System.Drawing.Size(46, 26);
            this.txtPageSize.TabIndex = 247;
            this.txtPageSize.Text = "100";
            this.txtPageSize.TextChanged += new System.EventHandler(this.txtPageSize_TextChanged);
            this.txtPageSize.Leave += new System.EventHandler(this.txtPageSize_Leave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(129, 12);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(135, 25);
            this.label3.TabIndex = 248;
            this.label3.Text = "条   当前页：";
            // 
            // lblCurrentPage
            // 
            this.lblCurrentPage.AutoSize = true;
            this.lblCurrentPage.ForeColor = System.Drawing.Color.Red;
            this.lblCurrentPage.Location = new System.Drawing.Point(249, 15);
            this.lblCurrentPage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCurrentPage.Name = "lblCurrentPage";
            this.lblCurrentPage.Size = new System.Drawing.Size(18, 20);
            this.lblCurrentPage.TabIndex = 249;
            this.lblCurrentPage.Text = "1";
            // 
            // lblSept
            // 
            this.lblSept.AutoSize = true;
            this.lblSept.Location = new System.Drawing.Point(266, 15);
            this.lblSept.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSept.Name = "lblSept";
            this.lblSept.Size = new System.Drawing.Size(13, 20);
            this.lblSept.TabIndex = 250;
            this.lblSept.Text = "/";
            // 
            // lblPageCount
            // 
            this.lblPageCount.AutoSize = true;
            this.lblPageCount.ForeColor = System.Drawing.Color.Red;
            this.lblPageCount.Location = new System.Drawing.Point(279, 15);
            this.lblPageCount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPageCount.Name = "lblPageCount";
            this.lblPageCount.Size = new System.Drawing.Size(18, 20);
            this.lblPageCount.TabIndex = 251;
            this.lblPageCount.Text = "1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(295, 13);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(33, 25);
            this.label1.TabIndex = 252;
            this.label1.Text = "共";
            // 
            // lblTotalCount
            // 
            this.lblTotalCount.AutoSize = true;
            this.lblTotalCount.ForeColor = System.Drawing.Color.Red;
            this.lblTotalCount.Location = new System.Drawing.Point(319, 16);
            this.lblTotalCount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTotalCount.Name = "lblTotalCount";
            this.lblTotalCount.Size = new System.Drawing.Size(36, 20);
            this.lblTotalCount.TabIndex = 253;
            this.lblTotalCount.Text = "100";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(393, 13);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 25);
            this.label5.TabIndex = 254;
            this.label5.Text = "条记录";
            // 
            // linkFirst
            // 
            this.linkFirst.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.linkFirst.Location = new System.Drawing.Point(468, 8);
            this.linkFirst.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.linkFirst.Name = "linkFirst";
            this.linkFirst.Size = new System.Drawing.Size(66, 43);
            this.linkFirst.TabIndex = 255;
            this.linkFirst.Text = "|<<";
            this.linkFirst.UseVisualStyleBackColor = true;
            this.linkFirst.Click += new System.EventHandler(this.linkFirst_Click);
            // 
            // linkPrevious
            // 
            this.linkPrevious.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.linkPrevious.Location = new System.Drawing.Point(561, 8);
            this.linkPrevious.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.linkPrevious.Name = "linkPrevious";
            this.linkPrevious.Size = new System.Drawing.Size(67, 43);
            this.linkPrevious.TabIndex = 256;
            this.linkPrevious.Text = "<";
            this.linkPrevious.UseVisualStyleBackColor = true;
            this.linkPrevious.Click += new System.EventHandler(this.linkPrevious_Click);
            // 
            // linkNext
            // 
            this.linkNext.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.linkNext.Location = new System.Drawing.Point(683, 8);
            this.linkNext.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.linkNext.Name = "linkNext";
            this.linkNext.Size = new System.Drawing.Size(63, 43);
            this.linkNext.TabIndex = 257;
            this.linkNext.Text = ">";
            this.linkNext.UseVisualStyleBackColor = true;
            this.linkNext.Click += new System.EventHandler(this.linkNext_Click_1);
            // 
            // linkLast
            // 
            this.linkLast.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.linkLast.Location = new System.Drawing.Point(789, 7);
            this.linkLast.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.linkLast.Name = "linkLast";
            this.linkLast.Size = new System.Drawing.Size(78, 44);
            this.linkLast.TabIndex = 258;
            this.linkLast.Text = ">>|";
            this.linkLast.UseVisualStyleBackColor = true;
            this.linkLast.Click += new System.EventHandler(this.linkLast_Click);//
            // 
            // txtPageNum
            // 
            this.txtPageNum.Location = new System.Drawing.Point(913, 9);
            this.txtPageNum.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPageNum.Name = "txtPageNum";
            this.txtPageNum.Size = new System.Drawing.Size(54, 26);
            this.txtPageNum.TabIndex = 259;
            this.txtPageNum.TextChanged += new System.EventHandler(this.txtPageNum_TextChanged);
            this.txtPageNum.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPageNum_KeyPress);
            // 
            // btnGo
            // 
            this.btnGo.BackColor = System.Drawing.Color.Transparent;
            this.btnGo.ForeColor = System.Drawing.Color.Black;
            this.btnGo.Location = new System.Drawing.Point(992, 4);
            this.btnGo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnGo.Name = "btnGo";
            this.btnGo.Size = new System.Drawing.Size(68, 47);
            this.btnGo.TabIndex = 260;
            this.btnGo.Text = "跳转";
            this.btnGo.UseVisualStyleBackColor = false;
            this.btnGo.Click += new System.EventHandler(this.btnGo_Click);
            // 
            // Paging
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnGo);
            this.Controls.Add(this.txtPageNum);
            this.Controls.Add(this.linkLast);
            this.Controls.Add(this.linkNext);
            this.Controls.Add(this.linkPrevious);
            this.Controls.Add(this.linkFirst);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblTotalCount);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblPageCount);
            this.Controls.Add(this.lblSept);
            this.Controls.Add(this.lblCurrentPage);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtPageSize);
            this.Controls.Add(this.label2);
            this.Name = "Paging";
            this.Size = new System.Drawing.Size(1096, 56);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPageSize;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblCurrentPage;
        private System.Windows.Forms.Label lblSept;
        private System.Windows.Forms.Label lblPageCount;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblTotalCount;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button linkFirst;
        private System.Windows.Forms.Button linkPrevious;
        private System.Windows.Forms.Button linkNext;
        private System.Windows.Forms.Button linkLast;
        private System.Windows.Forms.TextBox txtPageNum;
        private System.Windows.Forms.Button btnGo;
    }
}
