using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using static Radiant.MyCalss.公用;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ExcelAPP = Microsoft.Office.Interop.Excel.Application;

namespace Radiant.Myform
{
    public partial class 条形码 : Form
    {
        private ExcelAPP app;
        private bool isPiliang = false;
        private BarType barType;
        private WorkbookEvents_SheetActivateEventHandler sheetActivateHandler;
        private int startRow = 1;
        private Worksheet currentWorksheet;
        private bool isDisposed = false;

        // 构造函数
        public 条形码()
        {
            InitializeComponent();
        }

        // 第二个构造函数，接收Excel应用实例
        public 条形码(BarType bar, ExcelAPP excelAPP, bool b = false)
        {
            if (excelAPP == null)
            {
                throw new ArgumentNullException(nameof(excelAPP), "传入的 ExcelAPP 实例不能为 null");
            }

            app = excelAPP;
            barType = bar;
            isPiliang = b;

            InitializeComponent();
            InitializeExcelEventHandlers();
            InitializeFormSettings();
        }

        private void InitializeExcelEventHandlers()
        {
            try
            {
                Workbook activeWorkbook = app.ActiveWorkbook;
                if (activeWorkbook != null)
                {
                    sheetActivateHandler = new WorkbookEvents_SheetActivateEventHandler(Wb_SheetActivate);
                    activeWorkbook.SheetActivate += sheetActivateHandler;
                }

                UpdateCurrentWorksheet();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化Excel事件处理程序时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateCurrentWorksheet()
        {
            // 释放之前的工作表引用
            if (currentWorksheet != null)
            {
                Marshal.ReleaseComObject(currentWorksheet);
                currentWorksheet = null;
            }

            currentWorksheet = app?.ActiveSheet as Worksheet;
            if (currentWorksheet != null)
            {
                // 移除之前的事件处理程序，避免重复添加
                currentWorksheet.SelectionChange -= 事件改变;
                currentWorksheet.SelectionChange += 事件改变;

                // 窗体关闭时自动移除事件处理程序
                this.FormClosing += (sender, e) => currentWorksheet.SelectionChange -= 事件改变;
            }
        }

        private void InitializeFormSettings()
        {
            try
            {
                if (barType == BarType.QR_CODE)
                {
                    this.Text = "二维码生成器";
                    label1.Text = "二维码文本";
                }
                else if (barType == BarType.CODE_128)
                {
                    this.Text = "条形码生成器";
                    label1.Text = "条形码文本";
                    pictureBox1.Location = new System.Drawing.Point(50, 150);
                    pictureBox1.Size = new Size(400, 100);
                }
                else
                {
                    this.Text = "条码生成器";
                    label1.Text = "条码文本";
                    pictureBox1.Location = new System.Drawing.Point(50, 150);
                    pictureBox1.Size = new Size(400, 100);
                }

                if (isPiliang)
                {
                    SelectCom.Visible = isPiliang;
                    LastBut.Visible = isPiliang;
                    NextBut.Visible = isPiliang;
                    label2.Visible = isPiliang;
                    label3.Visible = isPiliang;

                    if (currentWorksheet != null)
                    {
                        GetHeader(currentWorksheet);
                    }
                }

                toolTip1.SetToolTip(RowText, "设置表头起始行");
                UpdateButtonStates();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化窗体设置时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Wb_SheetActivate(object Sh)
        {
            try
            {
                if (isDisposed) return;

                UpdateCurrentWorksheet();
                if (currentWorksheet != null)
                {
                    GetHeader(currentWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"工作表激活事件处理时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public BarType BarType
        {
            set { barType = value; }
        }

        private void 事件改变(Range Target)
        {
            try
            {
                if (Target == null || Target.Count > 1)
                {
                    textBox1.Text = "选择的单元格过多";
                    return;
                }

                string text = Target.Value2?.ToString() ?? "";
                if (!string.IsNullOrEmpty(text))
                {
                    textBox1.Text = text;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择单元格事件处理时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取表头
        /// </summary>
        private void GetHeader(Worksheet worksheet)
        {
            if (worksheet == null) return;

            try
            {
                SelectCom.Items.Clear();
                Header.Clear();

                Range usedRange = null;
                Range headerRange = null;

                try
                {
                    usedRange = worksheet.UsedRange;
                    int colCount = usedRange.Columns.Count;

                    if (colCount > 0)
                    {
                        headerRange = worksheet.Range[worksheet.Cells[startRow, 1], worksheet.Cells[startRow, colCount]];

                        foreach (Range cell in headerRange)
                        {
                            string value = cell.Value2?.ToString();

                            // 过滤空值
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                Header.Add(value);
                            }
                        }
                    }
                }
                finally
                {
                    if (headerRange != null) Marshal.ReleaseComObject(headerRange);
                    if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                }

                if (Header.Count > 0)
                {
                    SelectCom.Items.AddRange(Header.ToArray());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取表头时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<string> Header = new List<string>();

        /// <summary>
        /// 生成条码
        /// </summary>
        public void GenerateBarcode(string barcodeText, BarcodeFormat barcodeFormat, int width, int height)
        {
            if (string.IsNullOrEmpty(barcodeText))
            {
                MessageBox.Show("条码文本不能为空", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

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
                    Height = height,
                    Margin = 1
                };
            }
            else
            {
                writer.Options = new EncodingOptions
                {
                    Width = width,
                    Height = height,
                    Margin = 10
                };
            }

            try
            {
                var barcodeBitmap = writer.Write(barcodeText);
                pictureBox1.Image?.Dispose();
                pictureBox1.Image = barcodeBitmap;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成条码时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            生成条码(textBox1.Text.Trim());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!int.TryParse(RowText.Text, out startRow) || startRow < 1)
                {
                    MessageBox.Show("请输入有效的表头起始行号（大于等于1）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (currentWorksheet != null)
                {
                    GetHeader(currentWorksheet);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置表头行时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int 当前索引 = -1;
        private List<string> 数据 = new List<string>();

        private void 条形码_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                switch (e.KeyCode)
                {
                    case Keys.Enter:
                        生成条码(textBox1.Text.Trim());
                        e.Handled = true;
                        break;

                    case Keys.Up:
                        LastBar();
                        e.Handled = true;
                        break;

                    case Keys.Down:
                        NextBar();
                        e.Handled = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"按键处理时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                if (当前索引 <= 0) return;

                当前索引--;
                textBox1.Text = 数据[当前索引];

                if (ValidateBarcodeText(数据[当前索引]))
                {
                    生成条码(数据[当前索引]);
                }

                label3.Text = $"当前显示生成位置:{当前索引 + 1}/{数据.Count}";
                UpdateButtonStates();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"切换到上一个条码时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void NextBar()
        {
            try
            {
                // 修正条件判断：使用 > 而非 >=
                if (当前索引 >= 数据.Count - 1) return;

                当前索引++;
                textBox1.Text = 数据[当前索引];

                if (ValidateBarcodeText(数据[当前索引]))
                {
                    生成条码(数据[当前索引]);
                }

                label3.Text = $"当前显示生成位置:{当前索引 + 1}/{数据.Count}";
                UpdateButtonStates();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"切换到下一个条码时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("请输入条码文本", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (!ValidateBarcodeText(bartext))
                {
                    return;
                }

                if (Enum.TryParse(barType.ToString(), out BarcodeFormat barcodeFormat))
                {
                    GenerateBarcode(前缀Text.Text + bartext, barcodeFormat, pictureBox1.Width, pictureBox1.Height);
                }
                else
                {
                    MessageBox.Show("无效的条码类型", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成条码时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectCom_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                当前索引 = -1;
                pictureBox1.Image = null;
                数据.Clear();
                textBox1.Text = string.Empty;
                label3.Text = "当前显示位置为:";
                UpdateButtonStates();

                if (SelectCom.SelectedIndex < 0 || currentWorksheet == null) return;

                int colIndex = SelectCom.SelectedIndex + 1;
                Range usedRange = null;
                Range columnRange = null;

                try
                {
                    usedRange = currentWorksheet.UsedRange;
                    columnRange = usedRange.Columns[colIndex];

                    object[,] values = columnRange.Value as object[,];
                    if (values != null)
                    {
                        for (int i = startRow + 1; i <= values.GetLength(0); i++)
                        {
                            if (values[i, 1] != null)
                            {
                                数据.Add(values[i, 1].ToString());
                            }
                        }
                    }

                    if (数据.Count > 0)
                    {
                        当前索引 = 0;
                        textBox1.Text = 数据[0];

                        // 确保至少有一个有效数据
                        if (ValidateBarcodeText(数据[0]))
                        {
                            生成条码(数据[0]);
                        }

                        label3.Text = $"当前显示生成位置:{当前索引 + 1}/{数据.Count}";
                        UpdateButtonStates();
                    }
                    else
                    {
                        MessageBox.Show("所选列中没有有效数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                finally
                {
                    if (columnRange != null) Marshal.ReleaseComObject(columnRange);
                    if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载列数据时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                bool result = bool.Parse(Tag.ToString());
                this.Tag = !result;
                this.TopMost = !result;
                button3.Text = result ? "置顶" : "取消置顶";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置窗口置顶状态时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 验证条码文本是否有效
        /// </summary>
        private bool ValidateBarcodeText(string text)
        {
            try
            {
                if (string.IsNullOrEmpty(text))
                {
                    MessageBox.Show("条码文本不能为空", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 二维码最大字符限制
                if (barType == BarType.QR_CODE && text.Length > 7089)
                {
                    MessageBox.Show("二维码文本过长（最大支持7089个数字或4296个字符）", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 检查是否包含不支持的字符
                string invalidChars = @"<>:""/\|?*";
                if (text.Any(c => invalidChars.Contains(c)))
                {
                    MessageBox.Show("文本包含不支持的字符，请避免使用: <>:/\"|? *", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"验证条码文本时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                int.TryParse(跳转行text.Text, out int newStartRow);
                if (newStartRow >= 0 && newStartRow <= 数据.Count)
                {
                    string str = 数据[newStartRow - 1];
                    textBox1.Text = str;
                    生成条码(str);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}