using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;

[assembly: System.Reflection.Obfuscation(Feature = "Apply to type * when public: renaming; virtualization", Exclude = false)]

/// <summary>
/// 移除VBA工程密码（32位版本）
/// 原理:通过修改内存,欺骗VBE密码输入框,绕过密码验证
/// </summary>
public class ClassRemoveVbaPasswordBy32
{
    /// <summary>
    /// 虚拟保护 (32位版本)
    /// </summary>
    [DllImport("kernel32")]
    public static extern IntPtr VirtualProtect(IntPtr lpAddress, int dwSize, int flNewProtect, ref int lpflOldProtect);

    /// <summary>
    /// 获取导出函数地址
    /// </summary>
    [DllImport("kernel32")]
    public static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

    /// <summary>
    /// 获取模块句柄
    /// </summary>
    [DllImport("kernel32")]
    public static extern IntPtr GetModuleHandleA(string lpModuleName);

    /// <summary>
    /// 模式对话框函数
    /// </summary>
    [DllImport("user32.dll", EntryPoint = "DialogBoxParamA")]
    private static extern int DialogBoxParam(IntPtr hInstance, IntPtr pTemplateName,
        IntPtr hWndParent, IntPtr lpDialogFunc, IntPtr dwInitParam);

    /// <summary>
    /// Hook后的内存数据（32位版本6字节）
    /// </summary>
    public byte[] HookBytes = new byte[6];

    /// <summary>
    /// Hook前的内存数据
    /// </summary>
    public byte[] OriginBytes = new byte[6];

    /// <summary>
    /// 目标函数指针
    /// </summary>
    private IntPtr pFunc;

    /// <summary>
    /// Hook状态标志
    /// </summary>
    private bool Flag;

    /// <summary>
    /// 恢复原始内存数据
    /// </summary>
    private void RecoverBytes()
    {
        try
        {
            if (Flag)
            {
                Marshal.Copy(OriginBytes, 0, pFunc, 6);
            }
        }
        catch (Exception) { }
    }

    /// <summary>
    /// 对话框委托定义
    /// </summary>
    private delegate int DelegateDialogBoxParam(IntPtr hInstance, IntPtr pTemplateName,
        IntPtr hWndParent, IntPtr lpDialogFunc, IntPtr dwInitParam);

    /// <summary>
    /// 委托实例
    /// </summary>
    private static DelegateDialogBoxParam MsgBoxDelegate;

    /// <summary>
    /// VBE密码Hook核心函数（32位版本）
    /// </summary>
    private bool VbePasswordHook()
    {
        try
        {
            byte[] tmpBytes = new byte[6]; // 临时存储原始指令
            IntPtr p; // 委托指针
            int originProtect = 0; // 原始内存保护属性

            pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA");

            // 修改内存页为可写
            if (VirtualProtect(pFunc, 6, 0x40, ref originProtect) != IntPtr.Zero)
            {
                Marshal.Copy(pFunc, tmpBytes, 0, 6); // 读取原始指令

                // 检查是否已经Hook（首字节是否为PUSH DWORD）
                if (tmpBytes[0] != 0x68)
                {
                    Marshal.Copy(pFunc, OriginBytes, 0, 6); // 备份原始指令

                    // 创建委托并获取指针
                    MsgBoxDelegate = new DelegateDialogBoxParam(MyDialogBoxParam);
                    p = Marshal.GetFunctionPointerForDelegate(MsgBoxDelegate);

                    // 构建Hook指令：PUSH <函数地址> + RETN（6字节）
                    HookBytes[0] = 0x68; // PUSH DWORD
                    Buffer.BlockCopy(BitConverter.GetBytes(p.ToInt32()), 0, HookBytes, 1, 4); // 压入32位地址
                    HookBytes[5] = 0xC3; // RETN

                    // 写入Hook指令
                    Marshal.Copy(HookBytes, 0, pFunc, 6);

                    Flag = true;
                    return true;
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            throw new Exception("解除密码时出现异常:" + ex.Message);
        }
    }

    /// <summary>
    /// 自定义对话框处理函数
    /// </summary>
    private int MyDialogBoxParam(IntPtr hInstance, IntPtr pTemplateName,
        IntPtr hWndParent, IntPtr lpDialogFunc, IntPtr dwInitParam)
    {
        try
        {
            // 检测到VBE密码对话框（模板ID 4070）
            if (pTemplateName.ToInt32() == 4070)
            {
                return 1; // 直接返回成功代码
            }
            else
            {
                // 恢复原始函数并调用
                RecoverBytes();
                int result = DialogBoxParam(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam);
                VbePasswordHook(); // 重新应用Hook
                return result;
            }
        }
        catch (Exception ex)
        {
            throw new Exception("解除密码时出现异常:" + ex.Message);
        }
    }

    /// <summary>
    /// 执行密码解除操作
    /// </summary>
    public void ReleasePassword()
    {
        try
        {
            if (VbePasswordHook())
            {
                MessageBox.Show("VBA密码解除成功");
            }
            else
            {
                MessageBox.Show("VBA已经是解除状态");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }
}