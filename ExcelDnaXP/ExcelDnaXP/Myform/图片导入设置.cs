using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace Radiant.Myform
{
    public partial class 图片导入设置 : Form
    {
        private ExcelApp _excelApp;
        private ImageInsertOptions _options = new ImageInsertOptions();

        public 图片导入设置()
        {
            InitializeComponent();
            InitializeDefaults();
        }

        public 图片导入设置(ExcelApp application)
        {
            _excelApp = application;
            InitializeComponent();
            InitializeDefaults();
        }

        private void InitializeDefaults()
        {
            // 设置默认边距
            LeftText.Text = "5";
            TopText.Text = "5";

            // 设置默认尺寸
            WidthText.Text = "100";
            HeightText.Text = "100";

            // 初始化控件状态
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            // 根据填充样式更新控件状态
            bool isCustomSize = radioButton2.Checked;
            WidthText.Enabled = isCustomSize;
            HeightText.Enabled = isCustomSize;
        }

        // 填充样式枚举
        public enum FillStyle
        {
            NoFill,
            FillCell,
            CustomSize
        }

        // 填充方向枚举
        public enum FillDirection
        {
            Horizontal,
            Vertical
        }

        private bool InsertSingleImage(string imagePath, Range targetRange)
        {
            if (!File.Exists(imagePath))
            {
                MessageBox.Show($"图片文件不存在: {imagePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            double left = targetRange.Left + _options.LeftPadding;
            double top = targetRange.Top + _options.TopPadding;

            // 插入图片
            Shape shape = targetRange.Worksheet.Shapes.AddPicture(
                imagePath,
                _options.LinkToFile ? MsoTriState.msoTrue : MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                (float)left,
                (float)top,
                -1,
                -1
            );

            ApplyFillStyle(shape, targetRange);
            return true;
        }

        private void InsertImages(Range startRange)
        {
            if (_options.SelectedImagePaths == null || _options.SelectedImagePaths.Length == 0)
            {
                MessageBox.Show("请先选择图片文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                _excelApp.ScreenUpdating = false;
                _excelApp.EnableEvents = false;
                _excelApp.Calculation = XlCalculation.xlCalculationManual;

                Range currentRange = startRange;
                int successCount = 0;

                foreach (string imagePath in _options.SelectedImagePaths)
                {
                    if (InsertSingleImage(imagePath, currentRange))
                    {
                        successCount++;
                        // 移动到下一个单元格
                        currentRange = _options.FillDirection == FillDirection.Horizontal
                            ? currentRange.Offset[0, 1]
                            : currentRange.Offset[1, 0];
                    }
                }

                MessageBox.Show($"成功导入 {successCount} 张图片！", "成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理图片时出错: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _excelApp.ScreenUpdating = true;
                _excelApp.EnableEvents = true;
                _excelApp.Calculation = XlCalculation.xlCalculationAutomatic;
            }
        }

        private void ApplyFillStyle(Shape shape, Range targetRange)
        {
            switch (_options.FillStyle)
            {
                case FillStyle.FillCell:
                    // 填充单元格模式：调整图片大小以填充单元格
                    shape.Width = (float)(targetRange.Width - _options.LeftPadding * 2);
                    shape.Height = (float)(targetRange.Height - _options.TopPadding * 2);
                    break;

                case FillStyle.CustomSize:
                    // 自定义大小模式：调整单元格大小以适应图片尺寸
                    targetRange.ColumnWidth = (_options.CustomWidth - 12) / 7.0 + 1;
                    targetRange.RowHeight = _options.CustomHeight;

                    // 设置图片大小
                    shape.Width = (float)_options.CustomWidth;
                    shape.Height = (float)_options.CustomHeight;
                    break;
            }

            // 根据"压缩Y"选项设置图片位置
            shape.Placement = _options.LinkToFile
                ? XlPlacement.xlMoveAndSize  // 图片放置在单元格中（随单元格移动）
                : XlPlacement.xlFreeFloating; // 图片放置在单元格上方（自由浮动）
        }

        private void 导入_Click(object sender, EventArgs e)
        {
            try
            {
                if (_excelApp.Selection is Range rng)
                {
                    // 设置填充样式
                    if (radioButton2.Checked)
                    {
                        _options.FillStyle = FillStyle.CustomSize;
                    }
                    else if (radioButton3.Checked)
                    {
                        _options.FillStyle = FillStyle.FillCell;
                        // 填充单元格模式：设置默认边距
                        _options.LeftPadding = 5;
                        _options.TopPadding = 5;
                        LeftText.Text = "5";
                        TopText.Text = "5";
                    }
                    else
                    {
                        _options.FillStyle = FillStyle.NoFill;
                    }

                    // 设置填充方向
                    if (LeftToRingt.Checked)
                        _options.FillDirection = FillDirection.Horizontal;
                    else if (UpToDown.Checked)
                        _options.FillDirection = FillDirection.Vertical;

                    // 设置链接方式
                    _options.LinkToFile = 压缩Y.Checked;

                    // 设置边距和尺寸
                    _options.LeftPadding = Convert.ToDouble(LeftText.Text);
                    _options.TopPadding = Convert.ToDouble(TopText.Text);
                    _options.CustomWidth = Convert.ToDouble(WidthText.Text);
                    _options.CustomHeight = Convert.ToDouble(HeightText.Text);

                    InsertImages(rng);
                }
                else
                {
                    MessageBox.Show("请先选择单元格区域！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入图片时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 浏览文件But_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.gif;*.bmp|所有文件|*.*";
                    openFileDialog.Multiselect = true;
                    openFileDialog.Title = "选择图片文件";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        _options.SelectedImagePaths = openFileDialog.FileNames;
                        MessageBox.Show($"已选择 {_options.SelectedImagePaths.Length} 张图片。\n" +
                                      "请选择目标单元格，然后点击'导入'按钮。",
                                      "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 填充样式单选按钮改变事件
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }
    }

    public class ImageInsertOptions
    {
        public string[] SelectedImagePaths { get; set; }
        public 图片导入设置.FillStyle FillStyle { get; set; }
        public 图片导入设置.FillDirection FillDirection { get; set; }

        /// <summary>左边距(像素)</summary>
        public double LeftPadding { get; set; }

        /// <summary>上边距(像素)</summary>
        public double TopPadding { get; set; }

        /// <summary>自定义宽度(像素)</summary>
        public double CustomWidth { get; set; }

        /// <summary>自定义高度(像素)</summary>
        public double CustomHeight { get; set; }

        /// <summary>是否链接到文件</summary>
        public bool LinkToFile { get; set; }
    }
}