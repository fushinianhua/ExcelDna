using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZXing.QrCode;
using ZXing;

using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ZXing.Common;
using ExcelAPP = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using static Radiant.MyCalss.公用;

namespace Radiant.Myform
{
    public partial class 条形码 : Form
    {
        private ExcelAPP app;
        private bool IsPiliang = false;
        private BarType bartype;
        private WorkbookEvents_SheetActivateEventHandler sheetActivateHandler;

        // 第二个构造函数
        public 条形码(BarType bar, ExcelAPP excelAPP, bool b = false)
        {
            try
            {
                if (excelAPP == null)
                {
                    throw new ArgumentNullException(nameof(excelAPP), "传入的 ExcelAPP 实例不能为 null");
                }
                app = excelAPP;
                Workbook activeWorkbook = app.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    sheetActivateHandler = new WorkbookEvents_SheetActivateEventHandler(Wb_SheetActivate);
                    activeWorkbook.SheetActivate += sheetActivateHandler;
                }
                // 这里使用 null 条件运算符检查 app 是否为 null
                Worksheet worksheet = app?.ActiveSheet as Worksheet;
                if (worksheet != null)
                {
                    worksheet.SelectionChange += 事件改变;
                    this.FontChanged += (sende, e) =>
                    {
                        worksheet.SelectionChange -= 事件改变;
                    };
                }

                bartype = bar;
                IsPiliang = b;
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Wb_SheetActivate(object Sh)
        {
            Worksheet worksheet = Sh as Worksheet;
            if (worksheet != null)
            {
                GetHeader(worksheet);
            }
        }

        public BarType BarType
        {
            set { bartype = value; }
        }

        private void 事件改变(Range Target)
        {
            if (Target.Count > 1)
            {
                textBox1.Text = "选择的单元格过多";
                return;
            }
            string text = Target.Value2 + "";
            if (!string.IsNullOrEmpty(text))
            {
                textBox1.Text = text;
            }
        }

        private void 条形码_Load(object sender, EventArgs e)
        {
            Worksheet worksheet = (Worksheet)app.ActiveSheet;
            try
            {
                if (bartype == BarType.QR_CODE)
                {
                    this.Name = "二维码";
                    label1.Text = "二维码" + "文本";
                }
                else if (bartype == BarType.CODE_128)
                {
                    this.Name = "条形码";
                    label1.Text = "条形码" + "文本";
                    pictureBox1.Location = new System.Drawing.Point(50, 150);
                    pictureBox1.Size = new Size(400, 100);
                }
                else
                {
                    this.Name = "生成码";
                    label1.Text = "生成码" + "文本";
                    pictureBox1.Location = new System.Drawing.Point(50, 150);
                    pictureBox1.Size = new Size(400, 100);
                }
                if (IsPiliang)
                {
                    SelectCom.Visible = IsPiliang;
                    LastBut.Visible = IsPiliang;
                    NextBut.Visible = IsPiliang;
                    label2.Visible = IsPiliang;
                    label3.Visible = IsPiliang;

                    if (worksheet != null)
                    {
                        GetHeader(worksheet);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (worksheet != null)
                { Marshal.ReleaseComObject(worksheet); }
            }
        }

        private void GetHeader(Worksheet worksheet)
        {
            try
            {
                SelectCom.Items.Clear();
                Header.Clear();
                int col = 0, row = 0;
                row = worksheet.UsedRange.Rows.Count;
                col = worksheet.UsedRange.Columns.Count;
                if (row >= 1 && col > 1)
                {
                    Range rng = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, col]];
                    foreach (Range r in rng)
                    {
                        string va = r.Value2?.ToString();
                        Header.Add(va);
                    }
                    Marshal.ReleaseComObject(rng);
                }

                if (Header.Count > 0)
                {
                    SelectCom.Items.AddRange(Header.ToArray());
                }
            }
            catch (Exception)
            {
            }
        }

        private List<string> Header = new List<string>();
        /// <summary>
        /// 生成码
        /// </summary>
        /// <param name="barcodeText">条码文本</param>
        /// <param name="barcodeFormat">条码类型</param>
        /// <param name="width">高度</param>
        /// <param name="height">宽度</param>

        public void GenerateBarcode(string barcodeText, BarcodeFormat barcodeFormat, int width, int height)
        {
            var writer = new BarcodeWriter
            {
                Format = barcodeFormat
            };

            if (barcodeFormat == BarcodeFormat.QR_CODE)
            {
                writer.Options = new QrCodeEncodingOptions
                {
                    CharacterSet = "UTF-8",
                    Width = width,
                    Height = height
                };
            }
            else
            {
                writer.Options = new EncodingOptions
                {
                    Width = width,
                    Height = height
                };
            }

            try
            {
                var barcodeBitmap = writer.Write(barcodeText);
                pictureBox1.Image = barcodeBitmap;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成条码时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                生成条码(textBox1.Text.Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int 当前索引 = -1;

        private void 条形码_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                Keys key = e.KeyCode;
                switch (key)
                {
                    case Keys.Enter:
                        生成条码(textBox1.Text.Trim());
                        break;

                    case Keys.Up:

                        LastBar();
                        break;

                    case Keys.Down:

                        NextBar();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void UpdateButtonStates()
        {
            LastBut.Enabled = 当前索引 > 0;
            NextBut.Enabled = 当前索引 < 数据.Count - 1;
        }

        private void LastBar()
        {
            try
            {
                if (当前索引 <= 0)
                { return; }
                当前索引--;
                textBox1.Text = 数据[当前索引];
                生成条码(数据[当前索引]);
                label3.Text = $"当前显示生成位置:{当前索引 + 1}/{数据.Count}";

                UpdateButtonStates();
            }
            catch (Exception ex)
            {
                // 这里可以添加日志记录或者提示用户
                Console.WriteLine($"发生错误: {ex.Message}");
            }
        }

        private void NextBar()
        {
            try
            {
                if (当前索引 == 数据.Count - 1)
                {
                    return;
                }
                当前索引++;
                textBox1.Text = 数据[当前索引];
                生成条码(数据[当前索引]);
                label3.Text = $"当前显示生成位置:{当前索引 + 1}/{数据.Count}";

                UpdateButtonStates();
            }
            catch (Exception ex)
            {
                // 这里可以添加日志记录或者提示用户
                Console.WriteLine($"发生错误: {ex.Message}");
            }
        }

        private void LastBut_Click(object sender, EventArgs e)
        {
            LastBar();
        }

        private void NextBut_Click(object sender, EventArgs e)
        {
            NextBar();
        }

        private void 生成条码(string bartext)
        {
            try
            {
                if (string.IsNullOrEmpty(bartext))
                {
                    MessageBox.Show("请输入条码文本", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 假设 bartype 是一个 ComboBox，并且其 SelectedItem 可以转换为 BarcodeFormat
                if (Enum.TryParse(bartype.ToString(), out BarcodeFormat barcodeFormat))
                {
                    GenerateBarcode(bartext, barcodeFormat, pictureBox1.Width, pictureBox1.Height);
                }
                else
                {
                    MessageBox.Show("请选择有效的条码类型", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private List<string> 数据 = new List<string>();

        private void SelectCom_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                当前索引 = -1;
                pictureBox1.Image = null;
                数据.Clear();
                textBox1.Text = string.Empty;
                label3.Text = "当前显示位置为:";
                LastBut.Enabled = false;
                NextBut.Enabled = true;
                this.Focus();
                int col = SelectCom.SelectedIndex + 1;
                Worksheet sheet = (Worksheet)app.ActiveSheet;
                Range r = sheet.UsedRange;
                Range rng = r.Columns[col];
                object[,] values = rng.Value;
                for (int i = 2; i <= values.GetLength(0); i++)
                {
                    string value = values[i, 1].ToString();
                    数据.Add(value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}