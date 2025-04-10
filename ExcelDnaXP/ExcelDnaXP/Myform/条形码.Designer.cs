namespace ExcelDnaXP.Myform
{
    partial class 条形码
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
            this.label2 = new System.Windows.Forms.Label();
            this.SelectCom = new System.Windows.Forms.ComboBox();
            this.NextBut = new System.Windows.Forms.Button();
            this.LastBut = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(353, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 24);
            this.label2.TabIndex = 21;
            this.label2.Text = "单元格列";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label2.Visible = false;
            // 
            // SelectCom
            // 
            this.SelectCom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.SelectCom.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SelectCom.FormattingEnabled = true;
            this.SelectCom.Location = new System.Drawing.Point(438, 85);
            this.SelectCom.MaxDropDownItems = 15;
            this.SelectCom.Name = "SelectCom";
            this.SelectCom.Size = new System.Drawing.Size(107, 24);
            this.SelectCom.TabIndex = 20;
            this.SelectCom.TabStop = false;
            this.SelectCom.Visible = false;
            this.SelectCom.SelectedIndexChanged += new System.EventHandler(this.SelectCom_SelectedIndexChanged);
            // 
            // NextBut
            // 
            this.NextBut.Location = new System.Drawing.Point(453, 328);
            this.NextBut.Name = "NextBut";
            this.NextBut.Size = new System.Drawing.Size(76, 32);
            this.NextBut.TabIndex = 19;
            this.NextBut.Text = "下一个";
            this.NextBut.UseVisualStyleBackColor = true;
            this.NextBut.Visible = false;
            this.NextBut.Click += new System.EventHandler(this.NextBut_Click);
            // 
            // LastBut
            // 
            this.LastBut.Enabled = false;
            this.LastBut.Location = new System.Drawing.Point(356, 328);
            this.LastBut.Name = "LastBut";
            this.LastBut.Size = new System.Drawing.Size(76, 32);
            this.LastBut.TabIndex = 18;
            this.LastBut.Text = "上一个";
            this.LastBut.UseVisualStyleBackColor = true;
            this.LastBut.Visible = false;
            this.LastBut.Click += new System.EventHandler(this.LastBut_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(117, 30);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(306, 32);
            this.textBox1.TabIndex = 14;
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(15, 85);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(300, 300);
            this.pictureBox1.TabIndex = 17;
            this.pictureBox1.TabStop = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(453, 32);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(62, 32);
            this.button1.TabIndex = 16;
            this.button1.Text = "生成";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(12, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 32);
            this.label1.TabIndex = 15;
            this.label1.Text = "条形码文本";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(354, 364);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(175, 26);
            this.label3.TabIndex = 22;
            this.label3.Text = "当前显示生成位置:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // 条形码
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 399);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SelectCom);
            this.Controls.Add(this.NextBut);
            this.Controls.Add(this.LastBut);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.KeyPreview = true;
            this.MaximumSize = new System.Drawing.Size(586, 438);
            this.MinimumSize = new System.Drawing.Size(586, 438);
            this.Name = "条形码";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "条形码";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.条形码_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.条形码_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox SelectCom;
        private System.Windows.Forms.Button NextBut;
        private System.Windows.Forms.Button LastBut;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
    }
}