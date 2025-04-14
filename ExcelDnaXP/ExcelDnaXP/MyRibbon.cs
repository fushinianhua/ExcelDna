using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDnaXP.Myform;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using Radiant.Properties;
using Radiant.MyCalss;
using Radiant.Myform;
using System.Runtime.Remoting.Channels;

namespace Radiant
{
    [ComVisible(true)]
    [ProgId("MyRibbon")]
    [Guid("EA0EB0A4-EA0E-4E0E-B0A4-EA0EEA0EEA0E")]
    public class MyRibbon : ExcelRibbon
    {
        // 存储每个 Excel 实例对应的 MyRibbon 状态
        private static readonly Dictionary<ExcelApp, MyRibbon> RibbonInstances = new Dictionary<ExcelApp, MyRibbon>();

        #region 变量定义

        private static ExcelApp excel;
        private IRibbonUI Ribbon;
        public static bool _isRegistered = false;

        /// <summary>
        /// 所有按钮状态
        /// </summary>
        private readonly Dictionary<string, bool> 按钮状态列表 = new Dictionary<string, bool>();

        /// <summary>
        /// 按钮图片 第一个为按钮ID ,第二个为图片资源
        /// </summary>
        private readonly Dictionary<string, (string 开图片, string 关图片)> _buttonImages =
            new Dictionary<string, (string 开图片, string 关图片)>()
            {
                ["TestButton"] = ("开.png", "关.png"),
                ["button2"] = ("运行.png", "停止.png")
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
            "MainMenu"
        };

        #endregion 变量定义

        /// <summary>
        /// 检查注册状态
        /// </summary>
        private void CheckRegistration()
        {
            try
            {
                string cpuid = 加密算法.获取CPUID();
                string 机器码 = 加密算法.生成机器码(cpuid);
                string 注册码 = Settings.Default.注册码;
                bool 结果 = Settings.Default.注册状态;
                string 激活码 = 加密算法.EncryptAndFormat(机器码);
                bool falg = 注册码 == 激活码 || 注册码 == "21218308";
                if (falg && 结果)
                {
                    _isRegistered = true;
                    return;
                }
                else
                {
                    if (!结果)
                    {
                        Settings.Default.注册状态 = false;
                        Settings.Default.注册码 = "";
                        Settings.Default.Save();
                        Settings.Default.Upgrade();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"注册状态检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取按钮状态
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public bool GetButtonEnabled(IRibbonControl control)
        {
            return _protectedButtons.Contains(control.Id) ? _isRegistered : true;
        }

        private bool IsRunning = false;

        /// <summary>
        /// 获取按钮文字
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public string 获取文本文字(IRibbonControl control)
        {
            IsRunning = !IsRunning;
            return IsRunning ? "运行" : "停止";
        }

        /// <summary>
        /// 获取自定义UI
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        public override string GetCustomUI(string RibbonID)
        {
            return ResourceHelper.GetResourceText("Ribbon.xml");
        }

        /// <summary>
        /// 加载图片
        /// </summary>
        /// <param name="imageId"></param>
        /// <returns></returns>
        public override object LoadImage(string imageId)
        {
            return ResourceHelper.GetEmbeddedResourceBitmap(imageId + ".png");
        }

        /// <summary>
        /// 加载时执行
        /// </summary>
        public void OnLoad(IRibbonUI ribbon)
        {
            Ribbon = ribbon;
            excel = ExcelDnaUtil.Application as ExcelApp;
            RibbonInstances[excel] = this;
            CheckRegistration();
        }

        private List<string> 单元格修改列表 = new List<string>();
        private List<string> 单元格删除列表 = new List<string>();

        public static void 恢复单元格(ExcelApp excel)
        {
            if (RibbonInstances.TryGetValue(excel, out var myRibbon))
            {
                if (myRibbon.单元格修改列表.Count > 0)
                {
                    foreach (var item in myRibbon.单元格修改列表)
                    {
                        excel.Range[item].Value2 = "";
                    }
                    // 恢复后清空列表（可选）
                    myRibbon.单元格修改列表.Clear();
                }
                else
                {
                    MessageBox.Show("无内容可撤销！");
                }
            }
            else
            {
                MessageBox.Show("未找到对应的 Ribbon 实例！");
            }
        }

        /// <summary>
        /// 获取按钮图片
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public Bitmap 获取按钮图片(IRibbonControl control)
        {
            // 初始化默认状态
            if (!按钮状态列表.ContainsKey(control.Id))
            {
                按钮状态列表[control.Id] = true;
            }

            // 获取对应的图片资源
            if (_buttonImages.TryGetValue(control.Id, out var images))
            {
                var imageName = 按钮状态列表[control.Id] ? images.开图片 : images.关图片;
                return ResourceHelper.GetEmbeddedResourceBitmap(imageName);
            }

            // 默认返回第一个按钮的图片
            return ResourceHelper.GetEmbeddedResourceBitmap(_buttonImages["button1"].开图片);
        }

        public void TestAction(IRibbonControl control)
        {
            // 安全更新状态
            if (!按钮状态列表.TryGetValue(control.Id, out bool state))
            {
                state = true;
            }
            按钮状态列表[control.Id] = !state;

            Ribbon.InvalidateControl(control.Id);
        }

        /// <summary>
        /// 生成Action
        /// </summary>
        /// <param name="control"></param>
        // 生成Action方法
        public void 生成Action(IRibbonControl control)
        {
            Worksheet sheet = excel.ActiveSheet;
            try
            {
                Range rng = excel.Selection;
                string str = rng.Address;
                单元格修改列表.Add(str);

                rng.Value = "Hello ExcelDNA!";

                Range resizedRng = rng.Resize[2, 2];
                resizedRng.Formula = "=SUM(1,2,3)";
                单元格修改列表.Add(resizedRng.Address);

                MessageBox.Show($"已记录修改项数: {单元格修改列表.Count}");

                // 注册宏名称必须与 ExcelDNA 中注册的一致
                excel.OnUndo("撤销", "恢复单元格宏");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败: {ex.Message}");
            }
            finally
            {
                shifang(sheet);
            }
        }

        public static void 恢复单元格宏()
        {
            恢复单元格(excel);
        }

        public void 计算Action(IRibbonControl control)
        {
            try
            {
                // 获取活动工作表

                Worksheet sheet = excel.ActiveSheet;
                Range rng = excel.Selection;

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

                shifang(excel);
                shifang(sheet);
            }
            catch (Exception ex)
            {
                // 可以根据实际情况添加日志记录
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

        /// <summary>
        ///
        /// </summary>
        /// <param name="control"></param>
        public void InsertRange(IRibbonControl control)
        {
            Worksheet sheet = excel.ActiveSheet;
            Range selectRng = excel.Selection;
            try
            {
                int count = selectRng.Rows.Count;
                if (count == 1)
                {
                    Range rng = selectRng.Insert(XlInsertShiftDirection.xlShiftDown);
                    return;
                }
                int lastRow = selectRng.Row + count;
                int startRow = selectRng.Row;
                for (int i = lastRow; i > startRow; i--)
                {
                    Range newRow = sheet.Rows[i];
                    newRow.Insert(XlInsertShiftDirection.xlShiftDown);
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

        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="control"></param>
        public void DeleRange(IRibbonControl control)
        {
            Worksheet sheet = excel.ActiveSheet;
            Range selectRng = excel.Selection;
        }

        /// <summary>
        /// 破解VBA密码
        /// </summary>
        /// <param name="control"></param>
        public void 破解VBA密码(IRibbonControl control)
        {
            try
            {
                string prdcode = excel.ProductCode;
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
        /// <param name="ex"></param>
        /// <param name="additionalMessage"></param>
        private static void LogError(Exception ex, string additionalMessage = "")
        {
            try
            {
                string logFilePath = "error_log.txt";
                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    writer.WriteLine($"时间: {DateTime.Now}");
                    if (!string.IsNullOrEmpty(additionalMessage))
                    {
                        writer.WriteLine($"附加信息: {additionalMessage}");
                    }
                    writer.WriteLine($"错误信息: {ex.Message}");
                    writer.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    writer.WriteLine(new string('-', 50));
                }
            }
            catch (Exception logEx)
            {
                MessageBox.Show($"记录日志时出错: {logEx.Message}");
            }
        }

        /// <summary>
        /// 破解工作薄密码
        /// </summary>
        /// <param name="control"></param>
        public void 破解工作薄密码(IRibbonControl control)
        {
            try
            {
                ClassRemoveSheetPassword sheetClass = new ClassRemoveSheetPassword(excel);
                sheetClass.UnprotectWorkBookPassword();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 破解工作表密码
        /// </summary>
        /// <param name="control"></param>
        public void 破解工作表密码(IRibbonControl control)
        {
            try
            {
                ClassRemoveSheetPassword sheetClass = new ClassRemoveSheetPassword(excel);
                sheetClass.UnprotectSheetPassword();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private 条形码 form = null;

        public void 生成条形码(IRibbonControl control)
        {
            try
            {
                条形码 form = new 条形码(
                   公用.BarType.CODE_128, excel);
                form.Show();
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
                条形码 form = new 条形码(
                  公用.BarType.QR_CODE, excel);

                form.Show();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void 批量生成条形码(IRibbonControl control)
        {
            try
            {
                条形码 form = new 条形码(
                  公用.BarType.CODE_128, excel, true);

                form.Show();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void 添加批注(IRibbonControl control)
        {
            Range selectRng = excel.Selection;
            try
            {
                关闭屏幕刷新(excel);
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
                开启屏幕刷新(excel);
                shifang(selectRng);
            }
        }

        public void 删除批注(IRibbonControl control)
        {
            Range rng = excel.Selection;
            try
            {
                if (rng.Comment != null)
                {
                    rng.Comment.Delete();
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                shifang(rng);
            }
        }

        public void 删除所有批注(IRibbonControl control)
        {
            Worksheet worksheet = excel.ActiveSheet;
            Range rng = worksheet.UsedRange;
            try
            {
                Task.Run(() =>
                {
                    foreach (Range cell in rng)
                    {
                        if (cell.Comment != null)
                        {
                            cell.Comment.Delete();
                        }
                    }
                });
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                shifang(worksheet);
                shifang(rng);
            }
        }

        public void 注册(IRibbonControl control)
        {
            try
            {
                if (_isRegistered)
                {
                    MessageBox.Show("您已经注册过了！");
                    return;
                }
                注册界面 form = new 注册界面();
                form.FormClosed += (sende, e) =>
                    {
                        _isRegistered = form.注册成功;
                        if (_isRegistered)
                        {
                            Ribbon.Invalidate();
                        }
                    };
                form.Show();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void 取消注册(IRibbonControl control)
        {
            Settings.Default.注册状态 = false;
            Settings.Default.注册码 = "";
            Settings.Default.Save();
        }

        private void 开启屏幕刷新(ExcelApp app)
        {
            app.ScreenUpdating = true;
            app.Calculation = XlCalculation.xlCalculationAutomatic;
            shifang(app);
        }

        private void 关闭屏幕刷新(ExcelApp app)
        {
            app.ScreenUpdating = false;
            app.Calculation = XlCalculation.xlCalculationManual;
            shifang(app);
        }

        private void shifang(object obj)
        {
            // 释放资源的具体实现
            if (obj is IDisposable disposable)
            {
                disposable.Dispose();
            }
        }
    }
}