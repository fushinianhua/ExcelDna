using ExcelDna.Integration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using static Radiant.MyClass.GlobalEnum;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Image = System.Drawing.Image;
using Shape = Microsoft.Office.Interop.Excel.Shape;

namespace Radiant.Myform
{
    public partial class 图片导入设置 : Form
    {
        private readonly ExcelApp _excelApp;
        private const double Unit = 28.3465; // Excel单位转换系数

        /// <summary>
        /// 是否自定义图片大小
        /// </summary>
        private bool _isCustomSize = false;

        /// <summary>
        /// 用户选择的图片文件路径
        /// </summary>
        private readonly ImageInsertOptions _options = new ImageInsertOptions();

        /// <summary>
        /// 图片之间的间距(padding)
        /// </summary>
        private const double padding = 5.0; // 默认5点间距

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
            // 设置默认选项
            UpToDown.Checked = true;
            ImageSize.Checked = true;
            不改变位置大小.Checked = true;
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            // 根据填充样式更新控件状态
            bool isCustomSize = 自定义大小.Checked;
            WidthText.Enabled = isCustomSize;
            HeightText.Enabled = isCustomSize;
            TopText.Enabled = isCustomSize;
            LeftText.Enabled = isCustomSize;
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
                        TileLabel.Text = $"已选择 {_options.SelectedImagePaths.Length} 张图片";
                        //MessageBox.Show($"已选择 {_options.SelectedImagePaths.Length} 张图片。\n" +
                        //              "请选择目标单元格，然后点击'导入'按钮。",
                        //              "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 导入_Click(object sender, EventArgs e)
        {
            try
            {
                if (_excelApp.Selection is Range rng)
                {
                    获取参数();
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

                // 预调整行列大小（针对图片大小填充模式）
                if (_options.FillStyle == FillStyle.图片大小填充)
                {
                    AdjustRowAndColumnSizes(startRange);
                }

                int successCount = 0;
                int totalImages = _options.SelectedImagePaths.Length;

                // 根据填充模式选择不同的插入方法
                if (_options.FillStyle == FillStyle.不填充)
                {
                    successCount = InsertImagesNoFill(startRange);
                }
                else
                {
                    Range currentRange = startRange;
                    foreach (string imagePath in _options.SelectedImagePaths)
                    {
                        图片尺寸 尺寸 = 获取图片尺寸(imagePath, currentRange);
                        if (InsertSingleImage(imagePath, currentRange, 尺寸))
                        {
                            successCount++;
                            currentRange = GetNextRange(currentRange, _options.FillDirection);
                            if (currentRange == null) break;
                        }
                    }
                }

                MessageBox.Show($"成功导入 {successCount}/{totalImages} 张图片！", "成功",
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

        /// <summary>
        /// 插入图片（不填充模式）- 使用自定义大小并按方向排列
        /// </summary>
        private int InsertImagesNoFill(Range startRange)
        {
            int successCount = 0;
            double currentLeft = startRange.Left;
            double currentTop = startRange.Top;

            // 获取自定义大小
            double width = 100; // 默认值
            double height = 100;

            if (double.TryParse(WidthText.Text, out double w))
            {
                width = w * Unit;
            }

            if (double.TryParse(HeightText.Text, out double h))
            {
                height = h * Unit;
            }

            foreach (string imagePath in _options.SelectedImagePaths)
            {
                图片尺寸 尺寸 = new 图片尺寸
                {
                    左边距 = currentLeft,
                    上边距 = currentTop,
                    宽度 = width,
                    高度 = height
                };

                if (InsertSingleImage(imagePath, startRange, 尺寸))
                {
                    successCount++;

                    // 根据方向更新下一个图片的位置（添加间距）
                    if (_options.FillDirection == FillDirection.水平方向填充)
                    {
                        currentLeft += width + padding;
                    }
                    else
                    {
                        currentTop += height + padding;
                    }
                }
            }

            return successCount;
        }

        /// <summary>
        /// 预调整行列大小
        /// </summary>
        private void AdjustRowAndColumnSizes(Range startRange)
        {
            try
            {
                Worksheet worksheet = startRange.Worksheet;
                Range currentRange = startRange;

                foreach (string imagePath in _options.SelectedImagePaths)
                {
                    if (!File.Exists(imagePath)) continue;

                    using (Image image = Image.FromFile(imagePath))
                    {
                        // 直接使用图片尺寸调整行列
                        double rowHeight = ExcelUnitConverter.PixelsToRowHeight(image.Height);
                        double colWidth = ExcelUnitConverter.PixelsToColumnWidth(image.Width);
                        double suofang;
                        if (rowHeight > 409 || colWidth > 255)
                        {
                            suofang = Math.Min(409 / rowHeight, 255 / colWidth);
                            rowHeight *= suofang;
                            colWidth *= suofang;
                        }

                        // 调整行高
                        Range rowRange = (Range)worksheet.Rows[currentRange.Row];
                        rowRange.RowHeight = rowHeight;

                        // 调整列宽
                        Range colRange = (Range)worksheet.Columns[currentRange.Column];
                        colRange.ColumnWidth = colWidth;
                    }

                    currentRange = GetNextRange(currentRange, _options.FillDirection);
                    if (currentRange == null) break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"调整行列大小时出错: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 获取下一个目标单元格
        private Range GetNextRange(Range currentRange, FillDirection direction)
        {
            try
            {
                if (direction == FillDirection.水平方向填充)
                {
                    return currentRange.Offset[0, 1];
                }
                else
                {
                    return currentRange.Offset[1, 0];
                }
            }
            catch
            {
                return null;
            }
        }

        private 图片尺寸 获取图片尺寸(string imagePath, Range targetRange)
        {
            if (_isCustomSize)
                return GetCustomSize(targetRange);

            switch (_options.FillStyle)
            {
                case FillStyle.图片大小填充:
                    return GetFillByImageSize(targetRange);

                case FillStyle.不填充:
                    return GetNoFillSize();

                case FillStyle.单元格填充:
                default:
                    return GetFillByCellSize(imagePath, targetRange);
            }
        }

        private 图片尺寸 GetCustomSize(Range targetRange)
        {
            ValidateNumberInputs();

            return new 图片尺寸
            {
                左边距 = targetRange.Left + double.Parse(LeftText.Text),
                上边距 = targetRange.Top + double.Parse(TopText.Text),
                宽度 = double.Parse(WidthText.Text) * Unit,
                高度 = double.Parse(HeightText.Text) * Unit
            };
        }

        private void ValidateNumberInputs()
        {
            if (!double.TryParse(LeftText.Text, out _) ||
                !double.TryParse(TopText.Text, out _) ||
                !double.TryParse(WidthText.Text, out _) ||
                !double.TryParse(HeightText.Text, out _))
            {
                throw new ArgumentException("请输入有效的数字值");
            }
        }

        private 图片尺寸 GetFillByImageSize(Range targetRange)
        {
            return new 图片尺寸
            {
                左边距 = targetRange.Left + 2,
                上边距 = targetRange.Top + 2,
                宽度 = targetRange.Width - 4,
                高度 = targetRange.Height - 4
            };
        }

        /// <summary>
        /// 获取不填充模式的图片尺寸（使用自定义大小）
        /// </summary>
        private 图片尺寸 GetNoFillSize()
        {
            double width = 100; // 默认值
            double height = 100;

            if (double.TryParse(WidthText.Text, out double w))
            {
                width = w * Unit;
            }

            if (double.TryParse(HeightText.Text, out double h))
            {
                height = h * Unit;
            }

            return new 图片尺寸
            {
                左边距 = 0, // 位置由插入方法设置
                上边距 = 0,
                宽度 = width,
                高度 = height
            };
        }

        private 图片尺寸 GetFillByCellSize(string imagePath, Range targetRange)
        {
            double cellWidth = targetRange.Width - 4;
            double cellHeight = targetRange.Height - 4;

            if (!File.Exists(imagePath))
                return new 图片尺寸
                {
                    左边距 = targetRange.Left + 2,
                    上边距 = targetRange.Top + 2,
                    宽度 = cellWidth,
                    高度 = cellHeight
                };

            using (Image image = Image.FromFile(imagePath))
            {
                double imageRatio = (double)image.Width / image.Height;
                double cellRatio = cellWidth / cellHeight;

                double width, height;
                if (imageRatio > cellRatio)
                {
                    width = cellWidth;
                    height = cellWidth / imageRatio;
                }
                else
                {
                    height = cellHeight;
                    width = cellHeight * imageRatio;
                }

                return new 图片尺寸
                {
                    左边距 = targetRange.Left + (cellWidth - width) / 2 + 2,
                    上边距 = targetRange.Top + (cellHeight - height) / 2 + 2,
                    宽度 = width,
                    高度 = height
                };
            }
        }

        /// <summary>
        ///   插入单张图片
        /// </summary>
        private bool InsertSingleImage(string imagePath, Range targetRange, 图片尺寸 尺寸)
        {
            try
            {
                if (!File.Exists(imagePath))
                {
                    MessageBox.Show($"图片文件不存在: {Path.GetFileName(imagePath)}", "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                Shape shape = targetRange.Worksheet.Shapes.AddPicture(
                    imagePath,
                    _options.LinkToFile ? MsoTriState.msoTrue : MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    (float)尺寸.左边距,
                    (float)尺寸.上边距,
                    (float)尺寸.宽度,
                    (float)尺寸.高度
                );

                // 根据填充模式决定是否锁定宽高比
                shape.LockAspectRatio = _options.FillStyle == FillStyle.单元格填充 ||
                                        _options.FillStyle == FillStyle.不填充
                    ? MsoTriState.msoTrue
                    : MsoTriState.msoFalse;

                // 添加图片边框
                shape.Line.Weight = 1.5f;
                shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Silver);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入图片失败({Path.GetFileName(imagePath)}): {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void 获取参数()
        {
            try
            {
                // 1. 填充样式
                if (RangeSize.Checked)
                {
                    _options.FillStyle = FillStyle.单元格填充;
                }
                else if (ImageSize.Checked)
                {
                    _options.FillStyle = FillStyle.图片大小填充;
                }
                else if (自定义大小.Checked)
                {
                    _options.FillStyle = FillStyle.不填充;
                }
                else
                {
                    // 默认值
                    _options.FillStyle = FillStyle.单元格填充;
                }

                // 2. 填充方向
                _options.FillDirection = UpToDown.Checked ?
                    FillDirection.垂直方向填充 :
                    FillDirection.水平方向填充;

                // 3. 图片属性
                if (改变位置大小.Checked)
                {
                    _options.图片属性 = 属性枚举.改变大小和位置;
                }
                else if (不改变位置大小.Checked)
                {
                    _options.图片属性 = 属性枚举.不改变大小和位置;
                }
                else
                {
                    _options.图片属性 = 属性枚举.改变位置不改变大小;
                }

                _options.嵌入单元格 = false;
                _options.LinkToFile = false; // 默认不链接到文件
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取参数时出错: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 自定义大小_CheckedChanged(object sender, EventArgs e)
        {
            _isCustomSize = 自定义大小.Checked;
            UpdateControlStates();
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void TileLabel_Click(object sender, EventArgs e)
        {
        }
    }

    /// <summary>
    /// 图片尺寸结构
    /// </summary>
    public struct 图片尺寸
    {
        public double 左边距;
        public double 上边距;
        public double 宽度;
        public double 高度;
    }
}