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
        private ExcelApp _excelApp;
        private const double 单位 = 28.3465;

        /// <summary>
        /// 是否自定义图片大小
        /// </summary>
        private bool _IsCosutom = false;

        /// <summary>
        /// 用户选择的图片文件路径
        /// </summary>
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
            UpdateControlStates();
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
            try
            {
                if (_options.SelectedImagePaths == null || _options.SelectedImagePaths.Length == 0)
                {
                    MessageBox.Show("请先选择图片文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                _excelApp.ScreenUpdating = false;
                _excelApp.EnableEvents = false;
                _excelApp.Calculation = XlCalculation.xlCalculationManual;

                Range currentRange = startRange;
                int successCount = 0;
                foreach (string imagePath in _options.SelectedImagePaths)
                {
                    图片尺寸 尺寸 = 获取图片尺寸(imagePath, currentRange);
                    if (InsertSingleImage(imagePath, currentRange, 尺寸))
                    {
                        successCount++;
                        // 移动到下一个单元格
                        currentRange = _options.FillDirection == FillDirection.水平方向填充 ? currentRange.Offset[0, 1] : currentRange.Offset[1, 0];
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

        private 图片尺寸 获取图片尺寸(string imagePath, Range targetRange)
        {
            double top;
            double left;
            double height;
            double width;
            //  double left = targetRange.Left + _options.LeftPadding;
            //  double top = targetRange.Top + _options.TopPadding;
            if (_IsCosutom)
            {
                width = Convert.ToDouble(WidthText.Text) * 单位;
                height = Convert.ToDouble(HeightText.Text) * 单位;
                left = Convert.ToDouble(TopText.Text) * 单位;
                top = Convert.ToDouble(LeftText.Text) * 单位;
            }
            if (_options.FillStyle == FillStyle.图片大小填充)
            {
                (height, width) = GetImagePixels(imagePath);
                left = -1;
                top = -1;
            }
            else
            {
                width = targetRange.Width;
                height = targetRange.Height;
                left = -1;
                top = -1;
            }

            图片尺寸 尺寸 = new 图片尺寸
            {
                左边距 = left,
                上边距 = top,
                宽度 = height,
                高度 = width
            };

            return 尺寸;
            // 插入图片
        }

        private bool InsertSingleImage(string imagePath, Range targetRange, 图片尺寸 尺寸)
        {
            try
            {
                if (!File.Exists(imagePath))
                {
                    MessageBox.Show($"图片文件不存在: {imagePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                Shape shape = targetRange.Worksheet.Shapes.AddPicture(
                    imagePath,
                    _options.LinkToFile ? MsoTriState.msoTrue : MsoTriState.msoFalse,
                    MsoTriState.msoTrue,
                    (float)尺寸.左边距,
                    (float)尺寸.上边距,
                     Width = (int)尺寸.宽度,
                     Height = (int)尺寸.高度
                );
                shape.LockAspectRatio = MsoTriState.msoTrue;

                ApplyFillStyle(shape, targetRange);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void ApplyFillStyle(Shape shape, Range targetRange)
        {
        }

        // 获取图片像素尺寸
        public (int Width, int Height) GetImagePixels(string imagePath)
        {
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("图片文件不存在", imagePath);

            try
            {
                using (Image image = Image.FromFile(imagePath))
                {
                    return (image.Width, image.Height);
                }
            }
            catch (OutOfMemoryException ex)
            {
                // 处理非图片文件或损坏的图片
                throw new InvalidDataException("无法读取图片格式", ex);
            }
        }

        private void 获取参数()
        {
            try
            {
                if (radioButton2.Checked)
                {
                    _options.FillStyle = FillStyle.图片大小填充;
                }
                else
                    _options.FillStyle = FillStyle.单元格填充;
                if (UpToDown.Checked)
                {
                    _options.FillDirection = FillDirection.垂直方向填充;
                }
                else
                {
                    _options.FillDirection = FillDirection.水平方向填充;
                }
                if (改变位置大小.Checked)

                {
                    _options.图片属性 = 属性枚举.改变大小和位置;
                }
                else if (不改变位置大小.Checked)
                {
                    _options.图片属性 = 属性枚举.不改变大小和位置;
                }
                else { _options.图片属性 = 属性枚举.改变位置不改变大小; }
                _options.嵌入单元格 = false;
            }
            catch (Exception)
            {
                throw;
            }
        }

        // 填充样式单选按钮改变事件
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
        }

        private void UpdateControlStates()
        {
            // 根据填充样式更新控件状态
            bool isCustomSize = radioButton2.Checked;
            WidthText.Enabled = isCustomSize;
            HeightText.Enabled = isCustomSize;
            TopText.Enabled = isCustomSize;
            LeftText.Enabled = isCustomSize;
        }

        private void 自定义大小_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                _IsCosutom = 自定义大小.Checked;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}