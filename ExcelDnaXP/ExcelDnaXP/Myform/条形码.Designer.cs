namespace Radiant.Myform
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
            this.components = new System.ComponentModel.Container();
            this.label2 = new System.Windows.Forms.Label();
            this.SelectCom = new System.Windows.Forms.ComboBox();
            this.NextBut = new System.Windows.Forms.Button();
            this.LastBut = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.前缀Text = new System.Windows.Forms.TextBox();
            this.RowText = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.跳转行text = new System.Windows.Forms.TextBox();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(572, 38);
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
            this.SelectCom.Location = new System.Drawing.Point(657, 38);
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
            this.NextBut.Location = new System.Drawing.Point(633, 349);
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
            this.LastBut.Location = new System.Drawing.Point(536, 349);
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
            this.textBox1.Location = new System.Drawing.Point(227, 34);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(230, 32);
            this.textBox1.TabIndex = 14;
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(29, 85);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(428, 351);
            this.pictureBox1.TabIndex = 17;
            this.pictureBox1.TabStop = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(475, 36);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 32);
            this.button1.TabIndex = 16;
            this.button1.Text = "生成";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(132, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 32);
            this.label1.TabIndex = 15;
            this.label1.Text = "条形码文本";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(494, 410);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(171, 19);
            this.label3.TabIndex = 22;
            this.label3.Text = "当前显示生成位置:";
            this.label3.Visible = false;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(26, 35);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 30);
            this.label4.TabIndex = 23;
            this.label4.Text = "前缀";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // 前缀Text
            // 
            this.前缀Text.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.前缀Text.Location = new System.Drawing.Point(67, 34);
            this.前缀Text.Multiline = true;
            this.前缀Text.Name = "前缀Text";
            this.前缀Text.Size = new System.Drawing.Size(48, 32);
            this.前缀Text.TabIndex = 24;
            this.前缀Text.Text = "82";
            this.前缀Text.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // RowText
            // 
            this.RowText.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.RowText.Location = new System.Drawing.Point(575, 85);
            this.RowText.Multiline = true;
            this.RowText.Name = "RowText";
            this.RowText.Size = new System.Drawing.Size(48, 30);
            this.RowText.TabIndex = 25;
            this.RowText.Text = "1";
            this.RowText.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(657, 85);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(107, 32);
            this.button2.TabIndex = 26;
            this.button2.Text = "重新获取";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(657, 136);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(107, 26);
            this.button3.TabIndex = 27;
            this.button3.Tag = "false";
            this.button3.Text = "置顶窗体";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // 跳转行text
            // 
            this.跳转行text.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.跳转行text.Location = new System.Drawing.Point(536, 296);
            this.跳转行text.Multiline = true;
            this.跳转行text.Name = "跳转行text";
            this.跳转行text.Size = new System.Drawing.Size(61, 30);
            this.跳转行text.TabIndex = 28;
            this.跳转行text.Text = "1";
            this.跳转行text.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(633, 293);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(76, 32);
            this.button4.TabIndex = 29;
            this.button4.Text = "跳转";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // 条形码
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 456);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.跳转行text);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.RowText);
            this.Controls.Add(this.前缀Text);
            this.Controls.Add(this.label4);
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
            this.MaximumSize = new System.Drawing.Size(800, 600);
            this.MinimumSize = new System.Drawing.Size(586, 438);
            this.Name = "条形码";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Tag = "false";
            this.Text = "条形码";
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
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox 前缀Text;
        private System.Windows.Forms.TextBox RowText;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox 跳转行text;
        private System.Windows.Forms.Button button4;
    }
}