using ExcelDnaXP.MyClass;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int result = 1/*加密算法.ProcessRegistration("")*/;
            if (result == 1)
            {
                MyRibbon._isRegistered = true;
                MyRibbon.刷新();
                MessageBox.Show("注册成功");
                this.Close();
            }
        }
    }
}