using System.Collections.Generic;
using System.Drawing;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDnaTest;
using ExcelDnaXP.Myform;
using Microsoft.Office.Interop.Excel;

using Excel = Microsoft.Office.Interop.Excel;

using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelDnaXP.MyCalss;
using System.IO;
using System.Runtime.CompilerServices;
using System.Configuration;
using System.Security.Cryptography;
using System.Text;
using ExcelDemo.MyCalss;

namespace ExcelDnaXP
{
    [ComVisible(true)]
    [ProgId("MyRibbon")]
    [Guid("EA0EB0A4-EA0E-4E0E-B0A4-EA0EEA0EEA0E")]
    public class MyRibbon : ExcelRibbon
    {
        public MyRibbon()
        {
            CheckRegistration();
        }

        private static IRibbonUI Ribbon;
        public static bool _isRegistered = false;
        private const string RegKeyName = "ActivationCode";// 注册码的键名

        /// <summary>
        /// 检查注册状态
        /// </summary>
        private void CheckRegistration()
        {
            try
            {
                string jir = ConfigHelper.Appsettings.GetValue("注册状态");
                bool.TryParse(jir, out var result);
                if (result)
                {
                    _isRegistered = true;
                }
                var encryptedCode = ConfigurationManager.AppSettings[RegKeyName];
                if (!string.IsNullOrEmpty(encryptedCode))
                {
                    var machineCode = GenerateMachineCode();
                    var decryptedCode = UnprotectString(encryptedCode);
                    _isRegistered = decryptedCode == machineCode;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"注册状态检查失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 解密方法
        /// </summary>
        /// <param name="encryptedText"></param>
        /// <returns></returns>
        private static string UnprotectString(string encryptedText)
        {
            byte[] encryptedBytes = Convert.FromBase64String(encryptedText);
            byte[] plainBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.LocalMachine);
            return Encoding.UTF8.GetString(plainBytes);
        }

        // 生成机器特征码 ========================================
        private static string GenerateMachineCode()
        {
            var sha = SHA256.Create();
            var rawCode = $"{Environment.MachineName}{Environment.UserName}";
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(rawCode));
            return BitConverter.ToString(hash).Replace("-", "").Substring(0, 16);
        }

        public static void 刷新()
        {
            Ribbon.Invalidate();
        }

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
                ["TestButton"] = ("1.gif", "2.gif"),
                ["button2"] = ("Image3.png", "Image4.png")
            };

        /// <summary>
        /// 保护按钮
        /// </summary>

        private readonly List<string> _protectedButtons = new List<string> { "CalculateButton", "InsertButton" };

        public bool GetButtonEnabled(IRibbonControl control)
        {
            return _protectedButtons.Contains(control.Id) ? _isRegistered : true;
        }

        public override string GetCustomUI(string RibbonID)
        {
            return ResourceHelper.GetResourceText("Ribbon.xml");
        }

        public override object LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return ResourceHelper.GetEmbeddedResourceBitmap(imageId + ".png");
        }

        /// <summary>
        /// 加载时执行
        /// </summary>
        public void OnLoad(IRibbonUI ribbon) => Ribbon = ribbon;

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
        public void 生成Action(IRibbonControl control)
        {
            // 获取活动工作表
            dynamic excel = ExcelDnaUtil.Application;
            dynamic sheet = excel.ActiveSheet;
            try
            {
                // 向 A1 单元格写入数据
                sheet.Range["A1"].Value = "Hello ExcelDNA!";

                // 向 B2 单元格写入公式
                sheet.Range["B2"].Formula = "=SUM(1,2,3)";
            }
            catch (Exception)
            {
            }
            finally
            {
                shifang(excel);
                shifang(sheet);
            }
        }

        public void 计算Action(IRibbonControl control)
        {
            try
            {
                // 获取活动工作表
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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

        public void InsertRange(IRibbonControl control)
        {
            try
            {
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
                Worksheet sheet = excel.ActiveSheet;
                Range selectRng = excel.Selection;
                int lastRow = selectRng.Row + selectRng.Rows.Count;
                int startRow = selectRng.Row;
                for (int i = lastRow; i >= startRow; i -= 1)
                {
                    Range newRow = sheet.Rows[i];
                    newRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }
            }
            catch (Exception ex)
            {
                // 可以根据实际情况添加日志记录
                Console.WriteLine($"发生异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 破解VBA密码
        /// </summary>
        /// <param name="control"></param>
        public void 破解VBA密码(IRibbonControl control)
        {
            try
            {
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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

        public void 破解excel文件(IRibbonControl control)
        {
            try
            {
                // 创建文件选择对话框
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "Select an Excel Workbook";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    Task.Run(() => { CrackWorkbookPassword(filePath); });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private static void CrackWorkbookPassword(string filePath)
        {
            try
            {
                // 获取Excel应用程序对象
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;

                // 更全面的字符集，包含常见ASCII可打印字符和常见中文字符
                string charset = "";
                // 添加常见ASCII可打印字符
                for (int i = 32; i <= 126; i++)
                {
                    charset += (char)i;
                }
                // 添加常见中文字符，这里使用了汉字的Unicode范围
                for (int i = 0x4E00; i <= 0x9FA5; i++)
                {
                    charset += (char)i;
                }
                int maxLength = 20; // 最大密码长度，可根据实际情况调整

                // 生成密码并尝试破解
                for (int length = 1; length <= maxLength; length++)
                {
                    string[] passwords = GeneratePasswords(charset, length);
                    foreach (string password in passwords)
                    {
                        try
                        {
                            Workbook workbook = excel.Workbooks.Open(filePath, Password: password);
                            if (workbook != null)
                            {
                                // 密码正确
                                MessageBox.Show($"Password cracked: {password}");
                                workbook.Close(false);
                                return;
                            }
                        }
                        catch (COMException comEx)
                        {
                            // 记录密码错误尝试日志
                            LogError(comEx, $"尝试密码 {password} 失败。");
                            // 密码错误，继续尝试
                        }
                    }
                }
                MessageBox.Show("Password not found.");
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private static string[] GeneratePasswords(string charset, int length)
        {
            if (length == 1)
            {
                string[] result = new string[charset.Length];
                for (int i = 0; i < charset.Length; i++)
                {
                    result[i] = charset[i].ToString();
                }
                return result;
            }
            else
            {
                string[] prevPasswords = GeneratePasswords(charset, length - 1);
                string[] newPasswords = new string[prevPasswords.Length * charset.Length];
                int index = 0;
                foreach (string prevPassword in prevPasswords)
                {
                    for (int i = 0; i < charset.Length; i++)
                    {
                        newPasswords[index++] = prevPassword + charset[i];
                    }
                }
                return newPasswords;
            }
        }

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
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
                ClassRemoveSheetPassword sheetClass = new ClassRemoveSheetPassword(excel);
                sheetClass.UnprotectSheetPassword();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void 生成条形码(IRibbonControl control)
        {
            try
            {
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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
                ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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
            ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
            Range selectRng = excel.Selection;
            try
            {
                if (selectRng.Comment == null)
                {
                    selectRng.AddComment("批注");
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally

            {
                shifang(excel);
                shifang(selectRng);
            }
        }

        public void 删除批注(IRibbonControl control)
        {
            ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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
                shifang(excel);
                shifang(rng);
            }
        }

        public void 删除所有批注(IRibbonControl control)
        {
            ExcelApp excel = (ExcelApp)ExcelDnaUtil.Application;
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
                shifang(excel);
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
                }
                注册界面 form = new 注册界面();
                form.Show();
            }
            catch (Exception)
            {
                throw;
            }
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