using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDnaXP.Myform;
using Microsoft.Office.Interop.Excel;
using Microsoft.SqlServer.Server;
using Radiant.MyCalss;
using Radiant.MyClass;
using Radiant.Myform;
using Radiant.MyForm;
using Radiant.Properties;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Radiant
{
    [ComVisible(true)]
    [ProgId("MyRibbon")]
    [Guid("EA0EB0A4-EA0E-4E0E-B0A4-EA0EEA0EEA0E")]
    public class MyRibbon : ExcelRibbon
    {
        #region 变量定义

        // 实例级字段，每个Excel实例独立拥有
        private Application _excelApp;

        private IRibbonUI _ribbon;
        private ExcelAppEvents _appEvents;
        private bool _isRunning = false;
        private 条形码 _BarcodeForm = null;

        /// <summary>
        /// 按钮图片 第一个为按钮ID ,第二个为图片资源
        /// </summary>
        private readonly Dictionary<string, (string 开图片, string 关图片)> _buttonImages =
            new Dictionary<string, (string 开图片, string 关图片)>()
            {
                ["TestButton"] = ("开.png", "关.png"),
                //  ["button2"] = ("运行.png", "停止.png")
            };

        /// <summary>
        /// 保护按钮
        /// </summary>
        private readonly List<string> _protectedButtons = new List<string>
        {
            "CalculateButton",
            "批注",
            "InsertButton",
            "密码",
            "条码Menu",
            "MainMenu",
            "统计",
            "相同项"
        };

        /// <summary>
        /// 获取当前实例的状态管理对象
        /// </summary>
        private InstanceState CurrentInstanceState => InstanceManager.GetCurrentInstanceState();

        #endregion 变量定义

        /// <summary>
        /// 加载时执行
        /// </summary>
        public void OnLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            _excelApp = ExcelDnaUtil.Application as Application;

            // 注册工作簿事件，用于实例销毁时清理数据
            RegisterWorkbookEvents();

            //打开指定工作簿（仅在当前实例执行）
            //if (!string.IsNullOrEmpty("C:\\Users\\辛鹏\\Desktop\\test.xlsx"))
            //{
            //    _excelApp.Workbooks.Open("C:\\Users\\辛鹏\\Desktop\\test.xlsx");
            //}

            // 初始化状态
            CheckRegistration();
            //ShowSettingsPath();
        }

        /// <summary>
        /// 优化后的注册状态检查方法
        /// </summary>
        private void CheckRegistration()
        {
            try
            {
                // 确保加密算法已初始化
                加密算法.Initialize();

                string machineCode = 加密算法.注册码;
                string savedCode = Settings.Default.注册码;
                bool savedState = Settings.Default.激活状态;

                // 生成当前机器码对应的激活码
                string activationCode = 加密算法.EncryptAndFormat(加密算法.CPUID);

                // 验证逻辑：激活状态为True且注册码有效
                bool falg = savedState &&
                           (!string.IsNullOrEmpty(savedCode) &&
                            (savedCode == machineCode || savedCode == 加密算法.密钥));
                // 显示调试信息
                //MessageBox.Show($"机器码: {machineCode}\n" +
                //              $"激活码: {activationCode}\n" +
                //              $"注册码: {savedCode}\n" +
                //              $"密钥: {加密算法.密钥}\n" +
                //              $"结果: {falg}\n" +
                //              $"状态: {savedState}");

                // 更新实例状态
                CurrentInstanceState.IsRegistered = falg;
            }
            catch (Exception ex)
            {
                LogError(ex, "注册状态检查失败");
                //  MessageBox.Show($"注册状态检查失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取按钮状态
        /// </summary>
        public bool GetButtonEnabled(IRibbonControl control)
        {
            bool falg = _protectedButtons.Contains(control.Id) ? CurrentInstanceState.IsRegistered : true;
            return falg;
        }

        /// <summary>
        /// 获取按钮文字
        /// </summary>
        public string 获取文本文字(IRibbonControl control)
        {
            _isRunning = !_isRunning;
            return _isRunning ? "运行" : "停止";
        }

        /// <summary>
        /// 获取自定义UI
        /// </summary>
        public override string GetCustomUI(string RibbonID)
        {
            return ResourceHelper.GetResourceText("Ribbon.xml");
        }

        /// <summary>
        /// 加载图片
        /// </summary>
        public override object LoadImage(string imageId)
        {
            return ResourceHelper.GetEmbeddedResourceBitmap(imageId + ".png");
        }

        public void ShowSettingsPath()
        {
            try
            {
                // 方法1：获取.NET Application Settings的路径（若使用）
                string settingsPath = ConfigurationManager.OpenExeConfiguration(
                    ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;
                MessageBox.Show($"Settings路径: {settingsPath}");

                // 方法2：获取程序集目录（自定义配置文件可能在此）
                string assemblyPath = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location);
                MessageBox.Show($"程序集目录: {assemblyPath}");

                // 方法3：获取用户数据目录（自定义文件常见位置）
                string appDataPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData);
                MessageBox.Show($"用户数据目录: {appDataPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查找路径失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 注册工作簿事件，用于清理实例数据
        /// </summary>
        private void RegisterWorkbookEvents()
        {
            try
            {
                // 创建事件处理对象并保持引用
                _appEvents = new ExcelAppEvents(_excelApp, wb =>
                {
                    // 工作簿关闭时清理当前实例的数据
                    InstanceManager.CleanupCurrentInstanceState();
                });
            }
            catch (Exception ex)
            {
                LogError(ex, "注册工作簿事件失败");
            }
        }

        /// <summary>
        /// 获取按钮图片
        /// </summary>
        public Bitmap 获取按钮图片(IRibbonControl control)
        {
            // 从当前实例的状态中获取按钮状态
            if (!CurrentInstanceState.ButtonStates.TryGetValue(control.Id, out bool state))
            {
                state = true; // 默认启用
                CurrentInstanceState.ButtonStates[control.Id] = state;
            }

            // 获取对应的图片资源
            if (_buttonImages.TryGetValue(control.Id, out var images))
            {
                var imageName = state ? images.开图片 : images.关图片;
                return ResourceHelper.GetEmbeddedResourceBitmap(imageName);
            }

            // 默认返回第一个按钮的图片
            return ResourceHelper.GetEmbeddedResourceBitmap(_buttonImages["TestButton"].开图片);
        }

        public void TestAction(IRibbonControl control)
        {
            // 安全更新当前实例的按钮状态
            if (!CurrentInstanceState.ButtonStates.TryGetValue(control.Id, out bool state))
            {
                state = true;
            }
            CurrentInstanceState.ButtonStates[control.Id] = !state;

            _ribbon.InvalidateControl(control.Id);
        }

        /// <summary>
        /// 生成Action
        /// </summary>
        public void 生成Action(IRibbonControl control)
        {
            if (_excelApp == null) return;

            Worksheet sheet = _excelApp.ActiveSheet;
            try
            {
                Range rng = _excelApp.Selection;
                if (rng == null) return;
                Range resizedRng = rng.Resize[2, 2];
                resizedRng.Formula = "=SUM(1,2,3)";
                MessageBox.Show($"已记录修改项数: {rng.Count}");
            }
            catch (Exception ex)
            {
                LogError(ex, "生成操作失败");
                MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                shifang(sheet);
            }
        }

        public void 计算Action(IRibbonControl control)
        {
            try
            {
                Worksheet sheet = _excelApp.ActiveSheet;
                Range rng = _excelApp.Selection;

                if (rng == null) return;

                if (rng.Rows.Count > 1)
                {
                    MessageBox.Show("请选择一个单元格");
                    return;
                }

                int col = rng.Column;
                int startRow = GetStartRow(sheet, col);
                var (sendname1, sendname2) = GetNames(sheet, startRow, col);

                var (name1, name2, name1count, name2count) = ShowNameInputForm(sendname1, sendname2);

                int count = name1count + name2count;
                if (!string.IsNullOrEmpty(name1) && !string.IsNullOrEmpty(name2))
                {
                    object[,] value = GenerateValueArray(name1, name2, name1count, name2count);
                    Range range = sheet.Cells[startRow, col];
                    rng.Copy();
                    range.Resize[count, 1].PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    range.Resize[count, 1].Value2 = value;
                }

                shifang(_excelApp);
                shifang(sheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生异常: {ex.Message}");
            }
        }

        private int GetStartRow(Worksheet sheet, int col)
        {
            Range r = sheet.Cells[sheet.Rows.Count, col];
            return r.End[XlDirection.xlUp].Row + 1;
        }

        private (string, string) GetNames(Worksheet sheet, int startRow, int col)
        {
            string sendname1 = "YQ";
            string sendname2 = "YQ";
            if (startRow > 200)
            {
                Range valuerng = sheet.Range[sheet.Cells[startRow - 1, col], sheet.Cells[startRow - 200, col]];
                List<string> names = new List<string>();
                object[,] values = valuerng.Value2 as object[,];
                if (values != null)
                {
                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        string v = values[i, 1]?.ToString();
                        if (!string.IsNullOrEmpty(v) && !names.Contains(v))
                        {
                            names.Add(v);
                        }
                    }

                    if (names.Count >= 2)
                    {
                        sendname1 = names[0];
                        sendname2 = names[1];
                    }
                    else if (names.Count >= 1)
                    {
                        sendname1 = names[0];
                    }
                }
            }
            return (sendname1, sendname2);
        }

        private (string, string, int, int) ShowNameInputForm(string sendname1, string sendname2)
        {
            名字输入 form = new 名字输入();
            form.Load += (sender, e) =>
            {
                form.TextBox1Text = sendname1;
                form.TextBox2Text = sendname2;
            };
            string name1 = "";
            string name2 = "";
            int name1count = 0;
            int name2count = 0;

            form.FormClosed += (sender, e) =>
            {
                name1 = form.TextBox1Text;
                name2 = form.TextBox2Text;
                name1count = form.Textbox3Text;
                name2count = form.Textbox4Text;
            };

            form.ShowDialog();

            return (name1, name2, name1count, name2count);
        }

        private object[,] GenerateValueArray(string name1, string name2, int name1count, int name2count)
        {
            int count = name1count + name2count;
            object[,] value = new object[count, 1];
            for (int i = 0; i < count; i++)
            {
                if (i < name1count)
                {
                    value[i, 0] = name1;
                }
                else
                {
                    value[i, 0] = name2;
                }
            }
            return value;
        }

        public void 删除Actiond(IRibbonControl control)
        {
            MessageBox.Show("Hello!");
        }

        public void InsertRowRange(IRibbonControl control)//插入行
        {
            Worksheet sheet = _excelApp.ActiveSheet;
            Range selectRng = _excelApp.Selection;
            try
            {
                int Rowcount = selectRng.Rows.Count;
                if (selectRng == null || Rowcount == 1)
                {
                    return;
                }
                // 显示输入框获取插入行数
                object result = _excelApp.InputBox(
                    Prompt: "请输入插入行数",
                    Title: "插入行数",
                    Default: "1",
                    Type: 1);
                // 处理用户点击取消的情况
                if (result == null || result is bool v && !v || !(result is double))
                {
                    MessageBox.Show("输入无效，请输入数字", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return; // 用户取消操作
                }
                // 安全转换为整数

                double doubleValue = (double)result;
                // 四舍五入转换为整数
                int insertCount = (int)Math.Round(doubleValue, MidpointRounding.AwayFromZero);
                // 验证数值范围
                if (insertCount < 1 || insertCount > int.MaxValue)
                {
                    MessageBox.Show($"请输入1到{int.MaxValue}之间的整数", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int lastRow = selectRng.Row + Rowcount - 1;
                int startRow = selectRng.Row;
                for (int i = lastRow; i > startRow; i--)
                {
                    Range newRow = sheet.Rows[i];
                    // 一次性插入指定数量的空白行
                    newRow.Resize[insertCount].Insert(XlInsertShiftDirection.xlShiftDown);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                shifang(sheet);
                shifang(selectRng);
            }
        }

        public void InsertColRange(IRibbonControl control)
        {
            Worksheet worksheet = _excelApp.ActiveSheet;
            Range selectRng = _excelApp.Selection;

            try
            {
                // 检查选择区域是否有效
                if (selectRng == null || selectRng.Columns.Count < 2)
                {
                    MessageBox.Show("请至少选择两列进行操作", "提示",
                                   MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 获取用户输入的插入列数
                object result = _excelApp.InputBox(
                    Prompt: "请输入每列右侧插入的列数 (1-100)",
                    Title: "插入列",
                    Default: "1",
                    Type: 1);

                // 处理用户取消操作
                if (result == null || (result is bool boolResult && !boolResult))
                {
                    return; // 用户取消
                }

                // 增强输入验证
                if (!(result is double) && !double.TryParse(result.ToString(), out double temp))
                {
                    MessageBox.Show("输入无效，请输入一个数字", "错误",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                double doubleValue = result is double ? (double)result : double.Parse(result.ToString());

                // 添加合理的数值范围限制
                const int maxInsertCount = 100;
                if (doubleValue < 1 || doubleValue > maxInsertCount)
                {
                    MessageBox.Show($"请输入1到{maxInsertCount}之间的数字", "错误",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 四舍五入转换为整数
                int insertColumnCount = (int)Math.Round(doubleValue, MidpointRounding.AwayFromZero);

                // 关闭屏幕刷新
                关闭屏幕刷新(_excelApp);

                // 获取选择区域的列信息
                int selectedColsCount = selectRng.Columns.Count;
                int startColumn = selectRng.Column;
                int lastColumn = startColumn + selectedColsCount - 1;

                // 从最后一列开始处理，避免插入列后索引变化
                for (int i = lastColumn; i >= startColumn; i--)
                {
                    // 获取当前列的下一列作为插入点
                    Range insertPos = worksheet.Columns[i + 1];

                    try
                    {
                        // 正确的方式：插入整列区域
                        insertPos.Resize[ColumnSize: insertColumnCount].Insert(XlInsertShiftDirection.xlShiftToRight);
                    }
                    finally
                    {
                        // 确保释放临时创建的Range对象
                        shifang(insertPos);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}\n{ex.StackTrace}", "错误",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复Excel设置
                if (_excelApp != null)
                {
                    开启屏幕刷新(_excelApp);
                }

                // 释放COM对象
                shifang(selectRng);
                shifang(worksheet);
            }
        }//插入列

        public void 批量插入图片(IRibbonControl control)
        {
            try
            {
                图片导入设置 设置 = new 图片导入设置(_excelApp);
                设置.Show();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void 破解VBA密码(IRibbonControl control)
        {
            try
            {
                string prdcode = _excelApp.ProductCode;
                string[] pds = prdcode.Split('-');
                if (pds.Length > 4)
                {
                    if (pds[3].ToString().Equals("1000"))
                    {
                        ClassRemoveVbaPasswordBy64 VBEClass = new ClassRemoveVbaPasswordBy64();
                        VBEClass.ReleasePassword();
                    }
                    else
                    {
                        ClassRemoveVbaPasswordBy32 VBEClass = new ClassRemoveVbaPasswordBy32();
                        VBEClass.ReleasePassword();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}+移除vba密码异常", "Excel书世界");
            }
        }

        /// <summary>
        /// 记录错误信息
        /// </summary>
        /// <param name="ex">表示应用程序在运行时发生的错误信息</param>
        /// <param name="additionalMessage"></param>
        private static void LogError(Exception ex, string additionalMessage = "")
        {
            try
            {
                // 获取当前程序集的目录
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string assemblyDirectory = Path.GetDirectoryName(assemblyPath);

                // 确保日志文件路径在程序集同级目录
                string logFilePath = Path.Combine(assemblyDirectory, "error_log.txt");

                // 确保目录存在
                string logDirectory = Path.GetDirectoryName(logFilePath);
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    writer.WriteLine($"时间: {DateTime.Now}");
                    if (!string.IsNullOrEmpty(additionalMessage))
                    {
                        writer.WriteLine($"附加信息: {additionalMessage}");
                    }

                    if (ex != null)
                    {
                        writer.WriteLine($"错误信息: {ex.Message}");
                        writer.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    }
                    else
                    {
                        writer.WriteLine($"错误信息: 未提供异常对象");
                    }

                    writer.WriteLine(new string('-', 50));
                }
            }
            catch (Exception logEx)
            {
                MessageBox.Show($"记录日志时出错: {logEx.Message}");
            }
        }

        public void 破解工作薄密码(IRibbonControl control)
        {
            try
            {
                ClassRemoveSheetPassword sheetClass = new ClassRemoveSheetPassword(_excelApp);
                sheetClass.UnprotectWorkBookPassword();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void 破解工作表密码(IRibbonControl control)
        {
            try
            {
                ClassRemoveSheetPassword sheetClass = new ClassRemoveSheetPassword(_excelApp);
                sheetClass.UnprotectSheetPassword();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void 相同项对比(IRibbonControl control)
        {
            数据对比 dataCompare = new 数据对比(
                _excelApp);
            dataCompare.Show();
        }

        public void 生成条形码(IRibbonControl control)
        {
            try
            {
                ShowBarcodeForm(公用.BarType.CODE_128);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void 生成二维码(IRibbonControl control)
        {
            try
            {
                ShowBarcodeForm(公用.BarType.QR_CODE);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void 批量生成条形码(IRibbonControl control)//批量生成条形码
        {
            try
            {
                ShowBarcodeForm(公用.BarType.CODE_128, true);
            }
            catch (Exception)
            {
                throw;
            }
        }

        //显示条形码界面
        private void ShowBarcodeForm(公用.BarType barType, bool isBatch = false)
        {
            try
            {
                // 检查当前窗体是否存在，如果存在则关闭它
                if (_BarcodeForm != null && !_BarcodeForm.IsDisposed)
                {
                    // 可以选择提示用户，或直接关闭
                    _BarcodeForm.Close();
                }

                // 创建新窗体实例
                _BarcodeForm = new 条形码(barType, _excelApp, isBatch);

                // 设置窗体关闭时的事件处理，将引用置为null
                _BarcodeForm.FormClosed += (sender, e) => _BarcodeForm = null;

                // 显示窗体
                _BarcodeForm.Show();
            }
            catch (Exception ex)
            {
                LogError(ex, "显示条形码窗体失败");
                MessageBox.Show($"操作失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void 添加批注(IRibbonControl control) //批注
        {
            Range selectRng = _excelApp.Selection;
            try
            {
                关闭屏幕刷新(_excelApp);
                foreach (Range rng in selectRng)
                {
                    if (rng.Comment == null)
                    {
                        rng.AddComment("批注");
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                开启屏幕刷新(_excelApp);
                shifang(selectRng);
            }
        }

        public void 删除批注(IRibbonControl control)
        {
            Application excel = _excelApp;
            if (excel == null) return;

            Worksheet ws = excel.ActiveSheet;
            if (ws == null) return;

            Range selection = excel.Selection;
            if (selection == null) return;

            try
            {
                // 检查工作表保护
                if (ws.ProtectContents)
                {
                    MessageBox.Show("工作表受保护，无法删除批注", "警告",
                                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                foreach (Range cell in selection)
                {
                    if (cell.Comment != null)
                    {
                        cell.Comment.Delete();
                    }
                }
                // 显示操作结果
                MessageBox.Show("已删除选中区域内所有批注!", "操作完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除批注失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复Excel设置
                excel.ScreenUpdating = true;
                excel.EnableEvents = true;

                // 释放COM对象
                shifang(selection);
                shifang(ws);
            }
        }

        public void 删除所有批注(IRibbonControl control)
        {
            Application excel = _excelApp;
            if (excel == null) return;
            Worksheet worksheet = excel.ActiveSheet;
            if (worksheet == null) return;
            关闭屏幕刷新(excel);
            try
            {
                // 方法1: 使用SpecialCells快速定位批注（最推荐）
                try
                {
                    Range commentsRange = worksheet.Cells.SpecialCells(
                        XlCellType.xlCellTypeComments);

                    if (commentsRange != null)
                    {
                        // 一次性删除所有批注
                        commentsRange.ClearComments();
                        shifang(commentsRange);
                    }
                }
                catch (COMException ex) when (ex.ErrorCode == -2146827284) // 0x800A03EC
                {
                    // 找不到批注时忽略错误
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"删除批注时出错: {ex.Message}", "错误",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Range usedRange = worksheet.UsedRange; // 方法2: 备用方案（当SpecialCells失败时使用）
                if (usedRange != null)
                {
                    usedRange.ClearComments();   // 直接清除整个区域的批注（效率更高）
                }
            }
            finally
            {
                开启屏幕刷新(excel);  // 恢复屏幕刷新
                shifang(worksheet);         // 释放COM对象
            }
        }

        public void 数据统计(IRibbonControl control)
        {
            try
            {
                关闭屏幕刷新(_excelApp);
                Worksheet worksheet = _excelApp.Worksheets[2];//获取工作表
                if (worksheet == null || worksheet.Name != "MIB-OQC离职率统计") throw new Exception("工作表名称错误");
                Worksheet sheetdata = _excelApp.Worksheets[3];
                if (sheetdata == null || sheetdata.Name != "后道人员") throw new Exception("工作表名称错误");
                Range rng = worksheet.Range["C14:C33"];
                List<ItemData> itemDatas = new List<ItemData>();
                foreach (Range cell in rng.Cells)
                {
                    itemDatas.Add(new ItemData(cell.Value2.ToString(), cell.Row, 0));
                }
                DateTime dateTime;
                // 获取当前日期
                DateTime currentDate = DateTime.Now;

                // 计算前一天的日期
                DateTime previousDate = currentDate.AddDays(-1);

                if (previousDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    dateTime = currentDate.AddDays(-2);
                }
                else
                {
                    dateTime = previousDate;
                }

                int CurrentCol = 0;
                int Col = sheetdata.UsedRange.Columns.Count;
                int Row = sheetdata.UsedRange.Rows.Count;
                Range rng2 = sheetdata.Range[sheetdata.Cells[2, 1], sheetdata.Cells[2, Col]];
                Range rng3 = null;
                foreach (Range cell in rng2.Cells)
                {
                    if (IsDate(cell, dateTime, out CurrentCol))
                        break;
                }
                int 源列 = 0;
                Range rng4 = worksheet.Rows[4];
                foreach (Range cell in rng4.Cells)
                {
                    if (IsDate(cell, dateTime, out 源列))
                        break;
                }
                if (源列 <= 0)
                {
                    throw new Exception("没有找到日期列");
                }
                if (CurrentCol > 0)
                {
                    rng3 = sheetdata.Range[sheetdata.Cells[4, CurrentCol], sheetdata.Cells[Row, CurrentCol]];
                }
                else
                {
                    throw new Exception("没有找到日期列");
                }
                int CountAll = 0;
                int numbercount = 0;
                if (rng3 == null)
                {
                    throw new Exception("请选择数据范围");
                }
                List<Range> ranges = new List<Range>();
                foreach (Range cell in rng3.Cells)
                {
                    int number;
                    if (cell.Value2 == null) { continue; }
                    if (!int.TryParse(cell.Value.ToString(), out number))
                    {
                        ranges.Add(cell);
                    }
                    else
                    {
                        numbercount++;
                    }
                    CountAll++;
                }
                List<string> 正常离职列表 = new List<string>();
                List<string> 自离列表 = new List<string>();
                foreach (ItemData item in itemDatas)
                {
                    string str = item.Name;

                    foreach (Range cell in ranges)
                    {
                        string cellText = cell.Value.ToString();
                        switch (str)
                        {
                            case "曠工":
                                if (cellText.IndexOf("曠一", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    cellText.IndexOf("曠二", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    item.Count++;
                                }
                                break;

                            case "正常離職":
                                if (cellText.IndexOf("辦離職", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    item.Count++;
                                    int row = cell.Row;
                                    Range r = sheetdata.Cells[row, 3];
                                    正常离职列表.Add(r.Value);
                                    shifang(r);
                                }
                                break;

                            case "離職人數(自離)":
                                if (cellText.IndexOf("曠三", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    item.Count++;
                                    int row = cell.Row;
                                    Range r = sheetdata.Cells[row, 3];
                                    自离列表.Add(r.Value);
                                    shifang(r);
                                }
                                break;

                            default:
                                if (cellText.IndexOf(str, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    item.Count++;
                                }

                                break;
                        }
                    }
                }

                foreach (var item in itemDatas)
                {
                    Range r = worksheet.Cells[item.Row, 源列];
                    if (item.Count > 0)
                        r.Value2 = item.Count;
                    shifang(r);
                }
                // 将 List 内容用换行符连接
                string 正常离职列表字符串 = string.Join(Environment.NewLine, 正常离职列表);
                string 自离列表字符串 = string.Join(Environment.NewLine, 自离列表);
                Range range = null;
                if (正常离职列表.Count > 0)
                {
                    range = worksheet.Cells[32, 源列];
                    range.Comment?.Delete();

                    range.AddComment(正常离职列表字符串);
                }

                if (自离列表.Count > 0)
                {
                    range = worksheet.Cells[33, 源列];
                    range.Comment?.Delete();
                    range.AddComment(自离列表字符串);
                }

                shifang(range);
                // 确保CountAll不为零，防止除零错误
                if (CountAll <= 0)
                {
                    throw new Exception("统计总数为零，无法计算百分比");
                }

                // 定义要填充数据的单元格及其对应的值
                var dataToFill = new Dictionary<int, Func<string>>
                    {
                        { 6, () => CountAll.ToString() },
                        { 7, () => numbercount.ToString() },
                        { 8, () => (CountAll - numbercount).ToString() },
                        { 9, () => ((double)numbercount / CountAll).ToString("P2") },
                        { 10, () => ((double)(CountAll - numbercount) / CountAll).ToString("P2") }
                    };

                // 填充基础数据
                // 填充基础数据
                foreach (var item in dataToFill)
                {
                    worksheet.Cells[item.Key, 源列].Value = item.Value();
                }

                // 获取最后两项统计数据
                var value1 = itemDatas[itemDatas.Count - 2].Count;
                var value2 = itemDatas[itemDatas.Count - 1].Count;

                // 定义条件数据
                var conditionalData = new Dictionary<int, Func<string>>
                {
                    { 11, () => value1 > 0 ? ((double)value1 / CountAll).ToString("P2") : "0.00%"},
                    { 12, () => value2 > 0 ? ((double)value2 / CountAll).ToString("P2") : "0.00%"},
                    { 13, () => (value1 + value2) > 0 ? ((double)(value1 + value2) / CountAll).ToString("P2") :"0.00%" }
                };

                // 填充条件数据
                foreach (var item in conditionalData)
                {
                    string value = item.Value();
                    if (value != null)
                    {
                        worksheet.Cells[item.Key, 源列].Value = value;
                    }
                }

                开启屏幕刷新(_excelApp);
                MessageBox.Show("数据统计尚在优化");
            }
            catch (Exception ex)
            {
                开启屏幕刷新(_excelApp);
                MessageBox.Show(ex.Message);
            }
        }

        // 日期匹配方法
        private bool IsDate(Range rng, DateTime dateTime, out int CurrentCol)
        {
            CurrentCol = 0;

            // 跳过空单元格
            if (rng.Value2 == null)
            {
                return false;
            }

            try
            {
                DateTime cellDate;

                // 处理数字格式的日期
                if (rng.Value is double)
                {
                    cellDate = DateTime.FromOADate((double)rng.Value2);
                }
                // 处理字符串格式的日期
                else if (rng.Value is string)
                {
                    if (!DateTime.TryParse((string)rng.Value2, out cellDate))
                    {
                        return false;
                    }
                }
                // 处理DateTime对象
                else if (rng.Value is DateTime)
                {
                    cellDate = (DateTime)rng.Value;
                }
                else
                {
                    // 不是日期类型
                    return false;
                }

                // 比较日期部分
                if (cellDate.Date == dateTime.Date)
                {
                    // 保存匹配的列号
                    CurrentCol = rng.Column;
                    return true;
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }// 判断当前单元格是否为指定日期

        /// <summary>
        /// 注册功能
        /// </summary>
        public void 注册(IRibbonControl control)
        {
            try
            {
                if (CurrentInstanceState.IsRegistered)
                {
                    MessageBox.Show("您已经注册过了！", "注册提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                加密算法.Initialize();

                注册界面 form = new 注册界面
                {
                    机器码 = 加密算法.注册码,
                    Text = "Radiant插件注册"
                };

                form.FormClosed += (sender, e) =>
                {
                    if (form.激活状态)
                    {
                        // 保存用户输入的注册码
                        Settings.Default.注册码 = form.机器码;
                        Settings.Default.激活状态 = true;
                        Settings.Default.激活码 = form.激活码;
                        Settings.Default.Save();

                        // 更新实例状态
                        CurrentInstanceState.IsRegistered = true;

                        // 刷新Ribbon
                        if (_ribbon != null)
                        {
                            _ribbon.Invalidate();
                        }

                        MessageBox.Show("注册成功！", "注册结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (!string.IsNullOrEmpty(form.错误信息))
                    {
                        MessageBox.Show($"注册失败: {form.错误信息}", "注册结果", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                form.ShowDialog();
            }
            catch (Exception ex)
            {
                LogError(ex, "注册过程发生错误");
                MessageBox.Show($"注册过程发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 取消注册
        /// </summary>
        public void 取消注册(IRibbonControl control)
        {
            try
            {
                Settings.Default.激活状态 = false;
                Settings.Default.激活码 = "";
                Settings.Default.注册码 = "";
                Settings.Default.Save();

                CurrentInstanceState.IsRegistered = false;
                _ribbon.Invalidate();

                MessageBox.Show("已取消注册", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"取消注册失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 开启屏幕刷新(Application app)
        {
            app.ScreenUpdating = true;
            app.Calculation = XlCalculation.xlCalculationAutomatic;
        }

        private void 关闭屏幕刷新(Application app)
        {
            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;
        }

        private void shifang(object obj)
        {
            if (obj == null) return;

            try
            {
                if (obj is Range range)
                {
                    Marshal.ReleaseComObject(range);
                    range = null;
                }
                else if (obj is Worksheet sheet)
                {
                    shifang(sheet.UsedRange);
                    Marshal.ReleaseComObject(sheet);
                    sheet = null;
                }
                else if (obj is Application app)
                {
                    // 不建议在这里释放Application对象，可能导致Excel崩溃
                    // Marshal.ReleaseComObject(app);
                }
                else if (obj is IDisposable disposable)
                {
                    disposable.Dispose();
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "释放COM对象异常");
            }
        }
    }
}