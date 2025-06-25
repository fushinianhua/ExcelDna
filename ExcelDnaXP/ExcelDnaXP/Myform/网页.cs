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
    public partial class 网页 : Form
    {
        public string url { get; set; }

        public 网页()
        {
            InitializeComponent();
        }

        private void 网页_Load(object sender, EventArgs e)
        {
            webBrowser1.ScriptErrorsSuppressed = true; // 抑制脚本错误提示
            webBrowser1.Navigate(url);
        }
    }
}