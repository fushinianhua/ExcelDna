using System;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

[assembly: System.Reflection.Obfuscation(Feature = "Apply to type * when public: renaming; virtualization", Exclude = false)]

/// <summary>
/// 移除工作表保护密码
/// </summary>
public class ClassRemoveSheetPassword
{
    private Excel.Application XlAppObject;

    /// <summary>
    /// 初始化类
    /// </summary>
    /// <param name="xlapp"></param>
    public ClassRemoveSheetPassword(Excel.Application xlapp)
    {
        XlAppObject = xlapp;
    }

    /// <summary>
    /// 移除工作表密码
    /// </summary>
    public void UnprotectSheetPassword()
    {
        try
        {
            Excel.Worksheet actsheet = XlAppObject.ActiveSheet;
            if (!actsheet.ProtectContents)
            {
                MessageBox.Show("当前工作表未受保护！");

                return;
            }
            if (MessageBox.Show($"破解时如果要求输入密码,点击《取消》即可!{Environment.NewLine}一般情况下点击[1]次即可", "Excel书世界", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                object[] args1 = { true, true, true };
                string[] paramNames1 = { "DrawingObjects", "Contents", "AllowFiltering" };
                actsheet.GetType().InvokeMember("Protect", BindingFlags.InvokeMethod, null, actsheet, args1, null, null, paramNames1);

                object[] args2 = { false, true, true };
                string[] paramNames2 = { "DrawingObjects", "Contents", "AllowFiltering" };
                actsheet.GetType().InvokeMember("Protect", BindingFlags.InvokeMethod, null, actsheet, args2, null, null, paramNames2);

                object[] args3 = { true, true, true };
                string[] paramNames3 = { "DrawingObjects", "Contents", "AllowFiltering" };
                actsheet.GetType().InvokeMember("Protect", BindingFlags.InvokeMethod, null, actsheet, args3, null, null, paramNames3);

                object[] args4 = new object[0];
                actsheet.GetType().InvokeMember("UnProtect", BindingFlags.InvokeMethod, null, actsheet, args4);

                MessageBox.Show("密码清除完毕！");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    /// <summary>
    /// 移除工作簿密码
    /// </summary>
    public void UnprotectWorkBookPassword()
    {
        try
        {
            if (XlAppObject.ActiveWorkbook.ProtectStructure || XlAppObject.ActiveWorkbook.ProtectWindows)
            {
                if (MessageBox.Show("移除密码后将会产生一个《新的无密码保护的工作簿》,自行另存为即可!", "Excel书世界", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    XlAppObject.ActiveWorkbook.Sheets.Copy();
                    foreach (Excel.Worksheet sh in XlAppObject.ActiveWorkbook.Sheets)
                    {
                        sh.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    }
                    MessageBox.Show("工作簿密码清除完毕！");
                }
            }
            else
            {
                MessageBox.Show("当前工作簿未受保护！");

                return;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("密码清除异常！" + ex.Message);
        }
    }
}