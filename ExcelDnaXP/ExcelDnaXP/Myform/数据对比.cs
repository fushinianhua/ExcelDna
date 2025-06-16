using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Rectangle = System.Drawing.Rectangle;

namespace Radiant.MyForm
{
    public partial class 数据对比 : Form
    {
        // Excel应用程序实例
        private Application excelapp;

        // 标记颜色
        private Color? selectColor = Color.Yellow;

        // 可选颜色列表
        private List<Color> colors = new List<Color>();

        // 选择的两个对比区域
        private Range 区域一 = null;

        private Range 区域二 = null;

        // 对比结果区域
        private List<Range> 相同Rng = new List<Range>();

        private List<Range> 不同Rng = new List<Range>();

        // 对比结果键集合
        private HashSet<string> commonKeys;

        private HashSet<string> uniqueKeys1;
        private HashSet<string> uniqueKeys2;

        // 用于记录是否正在执行Excel操作
        private bool isProcessingExcel = false;

        public 数据对比(Application application)
        {
            InitializeComponent();
            InitializeColorComboBox();
            excelapp = application;
        }

        #region 窗体设计器生成的代码

        // 此处省略窗体设计器生成的代码，实际应用中需要通过设计器创建控件
        // 包括：pictureBox1, pictureBox2, 区域1Box, 区域2Box, 对比数据按钮,
        // 相同项按钮, 不同项按钮, 清除标识按钮, 导出相同项按钮, 导出不同项按钮,
        // 区域一Text, 区域二Text, 相同项Text 等控件

        #endregion 窗体设计器生成的代码

        /// <summary>
        /// 初始化颜色选择下拉框
        /// </summary>
        private void InitializeColorComboBox()
        {
            colorComboBox.SelectedIndexChanged += ColorComboBox_SelectedIndexChanged;
            colorComboBox.DrawItem += ColorComboBox_DrawItem;

            colors = new List<Color>
            {
                Color.Red,
                Color.Green,
                Color.Blue,
                Color.Yellow,
                Color.Orange,
                Color.Purple,
                Color.Orchid,
                Color.Pink,
                Color.PaleGreen,
                Color.Magenta
            };

            colorComboBox.DataSource = colors;
            colorComboBox.DisplayMember = "Name";
        }

        /// <summary>
        /// 颜色选择改变事件处理
        /// </summary>
        private void ColorComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = colorComboBox.SelectedIndex;
            if (index == colorComboBox.Items.Count - 1)
            {
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        selectColor = colorDialog.Color;
                    }
                }
            }
            else
            {
                selectColor = colors[index];
            }
        }

        /// <summary>
        /// 自定义绘制颜色下拉项
        /// </summary>
        private void ColorComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            var combo = sender as ComboBox;
            var color = (Color)combo.Items[e.Index];

            e.DrawBackground();

            Rectangle colorRect = new Rectangle(
                e.Bounds.X + 1,
                e.Bounds.Y + 1,
                combo.Width - 25,
                e.Bounds.Height - 4
            );

            if (e.Index == colors.Count - 1)
            {
                using (var brush = new SolidBrush(Color.White))
                {
                    e.Graphics.FillRectangle(brush, colorRect);
                }
                e.Graphics.DrawRectangle(Pens.Black, colorRect);
                e.Graphics.DrawString("更多颜色", new System.Drawing.Font("宋体", 8), Brushes.Black,
                    colorRect.X + 2, colorRect.Y + 2);
            }
            else
            {
                using (var brush = new SolidBrush(color))
                {
                    e.Graphics.FillRectangle(brush, colorRect);
                }
                e.Graphics.DrawRectangle(Pens.Black, colorRect);
            }
            e.DrawFocusRectangle();
        }

        [DllImport("user32.dll")]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// 选择第一个对比区域
        /// </summary>
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                ShowWaitCursor();
                BringExcelToFront();
                this.Hide();

                object result = excelapp.InputBox(
                    Prompt: "请选择单元格区域",
                    Title: "选择对比区域一",
                    Default: "",
                    Type: 8);

                if (result != null && result is Range selectedRange)
                {
                    区域一 = selectedRange;
                    区域1Box.Text = BuildFullAddress(selectedRange);
                }
                this.Show();
            }
            catch (Exception ex)
            {
                LogException("选择区域一失败", ex);
                MessageBox.Show($"选择区域一失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                HideWaitCursor();
            }
        }

        /// <summary>
        /// 选择第二个对比区域
        /// </summary>
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            try
            {
                ShowWaitCursor();
                BringExcelToFront();
                this.Hide();

                object result = excelapp.InputBox(
                    Prompt: "请选择单元格区域",
                    Title: "选择对比区域二",
                    Default: "",
                    Type: 8);

                if (result != null && result is Range selectedRange)
                {
                    区域二 = selectedRange;
                    区域2Box.Text = BuildFullAddress(selectedRange);
                }
                this.Show();
            }
            catch (Exception ex)
            {
                LogException("选择区域二失败", ex);
                MessageBox.Show($"选择区域二失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                HideWaitCursor();
            }
        }

        /// <summary>
        /// 构建完整的单元格地址
        /// </summary>
        private string BuildFullAddress(Range range)
        {
            if (range == null) return string.Empty;

            string workbookName = Path.GetFileName(range.Worksheet.Parent.FullName);
            string worksheetName = range.Worksheet.Name;
            string address = range.Address[XlReferenceStyle.xlA1].Replace("$", "");
            return $"[{workbookName}]{worksheetName}!{address}";
        }

        /// <summary>
        /// 对比数据按钮点击事件
        /// </summary>
        private void 对比数据_Click(object sender, EventArgs e)
        {
            try
            {
                ShowWaitCursor();
                EnableControls(false);

                // 清空之前的结果
                相同Rng.Clear();
                不同Rng.Clear();

                if (string.IsNullOrEmpty(区域1Box.Text) || string.IsNullOrEmpty(区域2Box.Text))
                {
                    MessageBox.Show("请先选择两个对比区域", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    EnableControls(true);
                    HideWaitCursor();
                    return;
                }

                // 获取数据
                object[,] data1 = 区域一.Value2 as object[,];
                object[,] data2 = 区域二.Value2 as object[,];

                if (data1 == null || data2 == null)
                {
                    MessageBox.Show("无法获取数据，请确保选择了有效区域", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    EnableControls(true);
                    HideWaitCursor();
                    return;
                }

                // 构建值字典
                var dict1 = BuildValueDictionary(data1, 区域一);
                var dict2 = BuildValueDictionary(data2, 区域二);

                // 计算对比结果
                commonKeys = new HashSet<string>(dict1.Keys.Intersect(dict2.Keys));
                uniqueKeys1 = new HashSet<string>(dict1.Keys.Except(dict2.Keys));
                uniqueKeys2 = new HashSet<string>(dict2.Keys.Except(dict1.Keys));

                // 收集相同项区域
                foreach (var key in commonKeys)
                {
                    相同Rng.AddRange(dict1[key]);
                    相同Rng.AddRange(dict2[key]);
                }

                // 收集不同项区域
                foreach (var key in uniqueKeys1)
                    不同Rng.AddRange(dict1[key]);

                foreach (var key in uniqueKeys2)
                    不同Rng.AddRange(dict2[key]);

                // 更新显示
                UpdateDisplay(commonKeys, uniqueKeys1, uniqueKeys2);

                LogInformation($"对比完成 - 相同项: {commonKeys.Count}, 区域一独有: {uniqueKeys1.Count}, 区域二独有: {uniqueKeys2.Count}");
            }
            catch (Exception ex)
            {
                LogException("对比数据失败", ex);
                MessageBox.Show($"对比数据失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                HideWaitCursor();
            }
        }

        /// <summary>
        /// 安全合并不同工作表的Range
        /// </summary>
        private Range SafeUnion(Range range1, Range range2)
        {
            if (range1 == null) return range2;
            if (range2 == null) return range1;

            try
            {
                // 记录当前活动工作表
                Worksheet originalSheet = excelapp.ActiveSheet;

                // 切换到第一个Range的工作表
                range1.Worksheet.Activate();

                // 执行Union操作
                Range result = excelapp.Union(range1, range2);

                // 切回原工作表
                originalSheet.Activate();

                return result;
            }
            catch (Exception ex)
            {
                LogWarning($"合并Range失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 合并多个Range的安全方法
        /// </summary>
        private Range MergeRanges(List<Range> ranges)
        {
            Range result = null;

            foreach (Range range in ranges)
            {
                if (range == null) continue;

                // 处理第一个Range
                if (result == null)
                {
                    result = range;
                    continue;
                }

                // 尝试合并
                try
                {
                    result = SafeUnion(result, range);
                }
                catch (Exception ex)
                {
                    LogWarning($"无法合并Range: {ex.Message}");
                }
            }

            return result;
        }

        /// <summary>
        /// 构建值到Range列表的字典
        /// </summary>
        private Dictionary<string, List<Range>> BuildValueDictionary(object[,] data, Range baseRange)
        {
            var dict = new Dictionary<string, List<Range>>();

            for (int row = 1; row <= data.GetLength(0); row++)
            {
                for (int col = 1; col <= data.GetLength(1); col++)
                {
                    var value = data[row, col];
                    var key = ConvertValueToKey(value);
                    var cell = baseRange.Cells[row, col];

                    if (!dict.ContainsKey(key))
                        dict[key] = new List<Range>();

                    dict[key].Add(cell);
                }
            }
            return dict;
        }

        /// <summary>
        /// 将值转换为用于比较的键
        /// </summary>
        private string ConvertValueToKey(object value)
        {
            if (value == null) return "∅";
            if (value is string str) return str.Trim();
            return Convert.ToString(value);
        }

        /// <summary>
        /// 更新界面显示
        /// </summary>
        private void UpdateDisplay(
            HashSet<string> commonKeys,
            HashSet<string> uniqueKeys1,
            HashSet<string> uniqueKeys2)
        {
            区域一Text.Text = string.Join(Environment.NewLine, uniqueKeys1);
            区域二Text.Text = string.Join(Environment.NewLine, uniqueKeys2);
            相同项Text.Text = string.Join(Environment.NewLine, commonKeys);

            // 更新按钮状态
            bool hasCommonItems = commonKeys.Count > 0;
            bool hasDifferentItems = uniqueKeys1.Count > 0 || uniqueKeys2.Count > 0;
            bool hasMarkedItems = 相同Rng.Count > 0 || 不同Rng.Count > 0;

            相同项.Enabled = hasCommonItems;
            导出相同项.Enabled = hasCommonItems;
            不同项.Enabled = hasDifferentItems;
            导出不同项.Enabled = hasDifferentItems;
            清除标识.Enabled = hasMarkedItems;
        }

        /// <summary>
        /// 标记相同项按钮点击事件
        /// </summary>
        private void 相同项_Click(object sender, EventArgs e)
        {
            ProcessMarking(相同Rng, "标记相同项");
        }

        /// <summary>
        /// 标记不同项按钮点击事件
        /// </summary>
        private void 不同项_Click(object sender, EventArgs e)
        {
            ProcessMarking(不同Rng, "标记不同项");
        }

        /// <summary>
        /// 处理标记操作的公共方法
        /// </summary>
        private void ProcessMarking(List<Range> ranges, string operationName)
        {
            try
            {
                if (!selectColor.HasValue || ranges.Count == 0) return;

                ShowWaitCursor();
                EnableControls(false);

                // 优化Excel性能
                excelapp.ScreenUpdating = false;
                excelapp.Calculation = XlCalculation.xlCalculationManual;

                // 按工作表分组处理Range
                var rangesByWorksheet = ranges.GroupBy(r => r.Worksheet);

                foreach (var group in rangesByWorksheet)
                {
                    Worksheet ws = group.Key;
                    Range combinedRange = null;

                    // 在同一工作表内合并Range
                    foreach (Range range in group)
                    {
                        if (combinedRange == null)
                            combinedRange = range;
                        else
                            combinedRange = SafeUnionSameSheet(combinedRange, range);
                    }

                    // 应用颜色标记
                    if (combinedRange != null)
                    {
                        ws.Activate();
                        combinedRange.Interior.Color = selectColor.Value.ToArgb();
                    }
                }

                // MessageBox.Show($"已成功{operationName}", "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogException($"{operationName}失败", ex);
                // MessageBox.Show($"{operationName}失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复Excel设置
                excelapp.ScreenUpdating = true;
                excelapp.Calculation = XlCalculation.xlCalculationAutomatic;

                EnableControls(true);
                HideWaitCursor();
            }
        }

        /// <summary>
        /// 安全合并同一工作表内的Range
        /// </summary>
        private Range SafeUnionSameSheet(Range range1, Range range2)
        {
            if (range1 == null) return range2;
            if (range2 == null) return range1;

            // 确保在同一工作表
            if (range1.Worksheet != range2.Worksheet)
                throw new ArgumentException("两个Range必须在同一工作表中");

            try
            {
                return excelapp.Union(range1, range2);
            }
            catch (Exception ex)
            {
                LogWarning($"合并Range失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 清除标记按钮点击事件
        /// </summary>
        private void 清除标识_Click(object sender, EventArgs e)
        {
            try
            {
                ShowWaitCursor();
                EnableControls(false);

                // 优化Excel性能
                excelapp.ScreenUpdating = false;
                excelapp.Calculation = XlCalculation.xlCalculationManual;

                // 合并相同和不同的Rng
                var allRanges = new List<Range>(相同Rng);
                allRanges.AddRange(不同Rng);

                var rangesByWorksheet = allRanges.GroupBy(r => r.Worksheet);

                foreach (var group in rangesByWorksheet)
                {
                    Worksheet ws = group.Key;
                    Range combinedRange = null;

                    foreach (Range range in group)
                    {
                        if (combinedRange == null)
                            combinedRange = range;
                        else
                            combinedRange = SafeUnionSameSheet(combinedRange, range);
                    }

                    if (combinedRange != null)
                    {
                        ws.Activate();
                        combinedRange.Interior.Color = XlRgbColor.rgbWhite; // 恢复为白色
                    }
                }

                MessageBox.Show("已清除所有标记", "操作完成", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 清空结果
                相同Rng.Clear();
                不同Rng.Clear();
                UpdateDisplay(commonKeys, uniqueKeys1, uniqueKeys2);
            }
            catch (Exception ex)
            {
                LogException("清除标记失败", ex);
                MessageBox.Show($"清除标记失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复Excel设置
                excelapp.ScreenUpdating = true;
                excelapp.Calculation = XlCalculation.xlCalculationAutomatic;

                EnableControls(true);
                HideWaitCursor();
            }
        }

        /// <summary>
        /// 导出相同项按钮点击事件
        /// </summary>
        private void 导出相同项_Click(object sender, EventArgs e)
        {
            ExportItems(相同Rng, "相同项");
        }

        /// <summary>
        /// 导出不同项按钮点击事件
        /// </summary>
        private void 导出不同项_Click(object sender, EventArgs e)
        {
            ExportItems(不同Rng, "不同项");
        }

        /// <summary>
        /// 导出项目的公共方法
        /// </summary>
        private void ExportItems(List<Range> ranges, string itemType)
        {
            if (ranges == null || ranges.Count == 0)
            {
                MessageBox.Show($"没有{itemType}可导出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                ShowWaitCursor();
                EnableControls(false);

                BringExcelToFront();
                this.Hide();

                Range targetRange = (Range)excelapp.InputBox(
                    Prompt: "请选择目标单元格",
                    Title: $"导出{itemType}",
                    Default: "",
                    Type: 8);

                if (targetRange != null)
                {
                    // 提取数据并去重
                    var values = ranges.Select(cell => cell.Value2)
                                      .Where(v => v != null)  // 过滤空值
                                      .Distinct()            // 去重
                                      .ToList();

                    // 转换为Excel所需的二维数组
                    object[,] data = ConvertToExcelArray(values);

                    // 导出数据
                    targetRange.Resize[data.GetLength(0), data.GetLength(1)].Value2 = data;

                    MessageBox.Show($"已导出 {values.Count} 个去重后的{itemType}",
                                   "导出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                this.Show();
            }
            catch (Exception ex)
            {
                LogException($"导出{itemType}失败", ex);
                MessageBox.Show($"导出{itemType}失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                HideWaitCursor();
            }
        }

        /// <summary>
        /// 转换一维数组为Excel可用的二维数组
        /// </summary>
        private object[,] ConvertToExcelArray(IEnumerable<object> values)
        {
            List<object> valueList = values.ToList();
            int count = valueList.Count;
            object[,] result = new object[count, 1];

            for (int i = 0; i < count; i++)
            {
                // 处理不同类型的值
                if (valueList[i] == null)
                    result[i, 0] = "∅";
                else if (valueList[i] is bool)
                    result[i, 0] = (bool)valueList[i] ? "TRUE" : "FALSE";
                else
                    result[i, 0] = valueList[i].ToString();
            }

            return result;
        }

        /// <summary>
        /// 窗体加载事件
        /// </summary>
        private void Form2_Load(object sender, EventArgs e)
        {
            // 初始化按钮状态
            相同项.Enabled = false;
            不同项.Enabled = false;
            清除标识.Enabled = false;
            导出不同项.Enabled = false;
            导出相同项.Enabled = false;

            // 启用键盘快捷键
            KeyPreview = true;
            KeyDown += Form_KeyDown;
        }

        /// <summary>
        /// 键盘按键事件处理
        /// </summary>
        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    this.Close();
                    break;

                case Keys.S:
                    相同项.PerformClick();
                    break;

                case Keys.D:
                    不同项.PerformClick();
                    break;

                case Keys.C:
                    清除标识.PerformClick();
                    break;

                case Keys.E:
                    导出相同项.PerformClick();
                    break;

                case Keys.F:
                    导出不同项.PerformClick();
                    break;
            }
            this.Focus();
        }

        /// <summary>
        /// 退出按钮点击事件
        /// </summary>
        private void 退出_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 窗体关闭事件 - 释放资源
        /// </summary>
        private void 数据对比_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 释放COM对象资源
            ReleaseComObject(区域一);
            ReleaseComObject(区域二);
            // 注意：Excel应用程序实例不应在此处释放，因为它是由外部传入的
        }

        /// <summary>
        /// 显示等待光标
        /// </summary>
        private void ShowWaitCursor()
        {
            this.Cursor = Cursors.WaitCursor;
            isProcessingExcel = true;
        }

        /// <summary>
        /// 隐藏等待光标
        /// </summary>
        private void HideWaitCursor()
        {
            this.Cursor = Cursors.Default;
            isProcessingExcel = false;
        }

        /// <summary>
        /// 启用或禁用控件
        /// </summary>
        private void EnableControls(bool enable)
        {
            pictureBox1.Enabled = enable;
            pictureBox2.Enabled = enable;
            对比数据.Enabled = enable;
            相同项.Enabled = enable && 相同项.Enabled;
            不同项.Enabled = enable && 不同项.Enabled;
            清除标识.Enabled = enable && 清除标识.Enabled;
            导出相同项.Enabled = enable && 导出相同项.Enabled;
            导出不同项.Enabled = enable && 导出不同项.Enabled;
        }

        /// <summary>
        /// 将Excel窗口置于前台
        /// </summary>
        private void BringExcelToFront()
        {
            if (excelapp != null)
            {
                IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
                SetForegroundWindow(excelHandle);
                SwitchToThisWindow(excelHandle, true);
            }
        }

        /// <summary>
        /// 释放COM对象资源
        /// </summary>
        private void ReleaseComObject(object comObject)
        {
            if (comObject == null) return;

            try
            {
                Marshal.ReleaseComObject(comObject);
                comObject = null;
            }
            catch (Exception ex)
            {
                LogWarning($"释放COM对象失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 记录信息日志
        /// </summary>
        private void LogInformation(string message)
        {
            Debug.WriteLine($"[INFO] {message}");
        }

        /// <summary>
        /// 记录警告日志
        /// </summary>
        private void LogWarning(string message)
        {
            Debug.WriteLine($"[WARNING] {message}");
        }

        /// <summary>
        /// 记录异常日志
        /// </summary>
        private void LogException(string message, Exception ex)
        {
            Debug.WriteLine($"[ERROR] {message}: {ex.Message}");
            Debug.WriteLine($"[STACK TRACE] {ex.StackTrace}");
        }
    }
}