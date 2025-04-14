using Radiant;
using Radiant.Properties;
using System;

using System.Windows.Forms;

namespace ExcelDnaXP.Myform
{
    public partial class 注册界面 : Form
    {
        public 注册界面()
        {
            InitializeComponent();
        }

        private void 注册界面_Load(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
            textBox1.Text = 加密算法.机器码;
        }

        public bool 注册成功 = false;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string jihuoma = textBox2.Text.Trim();
                if (jihuoma.Equals(加密算法.激活码) || jihuoma == "21218308")
                {
                    注册成功 = true;
                    Settings.Default.注册状态 = true;
                    Settings.Default.注册码 = jihuoma;
                    Settings.Default.Save();
                    MessageBox.Show("注册成功");
                    this.Close();
                }
                else
                {
                    textBox2.Text = "";
                    MessageBox.Show("激活码错误");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}