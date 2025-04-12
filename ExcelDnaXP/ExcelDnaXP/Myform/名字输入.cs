using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Radiant.Myform
{
    public partial class 名字输入 : Form
    {
        private bool IsOk = false;

        public 名字输入()
        {
            InitializeComponent();
        }

        private void 名字输入_Load(object sender, EventArgs e)
        {
            textBox1.Text = "张三";
            textBox2.Text = "李四";
        }

        // 公共属性，用于获取 TextBox 的文本
        public string TextBox1Text
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }

        public string TextBox2Text
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        public int Textbox3Text
        {
            get { return int.Parse(textBox3.Text.Trim()); }
            set { textBox3.Text = value.ToString(); }
        }

        public int Textbox4Text
        {
            get { return int.Parse(textBox4.Text.Trim()); }
            set { textBox4.Text = value.ToString(); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int c;
            string text = textBox3.Text;
            string text2 = textBox4.Text;

            bool flag = int.TryParse(text, out c);
            bool flag2 = int.TryParse(text2, out c);

            if (flag && flag2)
            {
                IsOk = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("请输入数字");
            }
        }

        private void 名字输入_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!IsOk)
            {
                e.Cancel = true;
            }
        }

        private void 名字输入_Load_1(object sender, EventArgs e)
        {
        }
    }
}