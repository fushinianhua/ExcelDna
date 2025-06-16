using Radiant;
using System;
using System.Windows.Forms;

namespace ExcelDnaXP.Myform
{
    public partial class 注册界面 : Form
    {
        public string 机器码 { get; set; }
        public string 注册码 { get; set; }
        public string 激活码 { get; private set; }
        public bool 激活状态 { get; private set; }
        public string 错误信息 { get; private set; }

        public 注册界面()
        {
            InitializeComponent();
        }

        private void 注册界面_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(机器码))
            {
                textBox1.Text = 机器码;
            }
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            激活验证();
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // 执行你的方法
                激活验证();

                // 取消默认行为（阻止换行）
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }

        private void 激活验证()
        {
            try
            {
                激活码 = textBox2.Text.Trim();

                try
                {
                    if (加密算法.ValidateRegistration(激活码))
                    {
                        激活状态 = true;
                        this.Close();
                    }
                    else
                    {
                        错误信息 = "注册码无效，请检查输入";
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    错误信息 = $"注册码验证异常: {ex.Message}";
                    this.Close();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}