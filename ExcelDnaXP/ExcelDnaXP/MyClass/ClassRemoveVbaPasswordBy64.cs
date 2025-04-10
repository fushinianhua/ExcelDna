using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: System.Reflection.Obfuscation(Feature = "Apply to type * when public: renaming; virtualization", Exclude = false)]

/// <summary>
/// 移除VBA工程密码
/// 原理:通过修改内存,欺骗VBE密码输入框,绕过密码验证
/// </summary>
public class ClassRemoveVbaPasswordBy64
{
    /// <summary>
    /// 虚拟保护
    /// </summary>
    /// <param name="lpAddress">内存起始地址</param>
    /// <param name="dwSize">内存区域大小</param>
    /// <param name="flNewProtect">指示可读可写可执行</param>
    /// <param name="lpflOldProtect">内存原始属性类型保存地址</param>
    /// <returns></returns>
    [DllImport("kernel32")]
    public static extern IntPtr VirtualProtect(IntPtr lpAddress, int dwSize, int flNewProtect, ref int lpflOldProtect);

    /// <summary>
    /// 导出函数的地址
    /// </summary>
    /// <param name="hModule">模块句柄</param>
    /// <param name="procName">函数名</param>
    /// <returns></returns>
    [DllImport("kernel32")]
    public static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

    /// <summary>
    /// 获取一个应用程序或动态链接库的模块句柄
    /// </summary>
    /// <param name="lpModuleName">模块名</param>
    /// <returns></returns>
    [DllImport("kernel32")]
    public static extern IntPtr GetModuleHandleA(string lpModuleName);

    /// <summary>
    /// 模式对话框
    /// </summary>
    /// <param name="hInstance">对话框模版所在模块的句柄</param>
    /// <param name="pTemplateName">对话框模版</param>
    /// <param name="hWndParent">拥有对话框窗口的句柄</param>
    /// <param name="lpDialogFunc">对话框消息处理函数</param>
    /// <param name="dwInitParam">传递给对话框过程的消息</param>
    /// <returns></returns>
    [DllImport("user32.dll", EntryPoint = "DialogBoxParamA")]
    private static extern int DialogBoxParam(IntPtr hInstance, IntPtr pTemplateName, IntPtr hWndParent, IntPtr lpDialogFunc, IntPtr dwInitParam);

    /// <summary>
    /// Hook之后的内存数据
    /// </summary>
    public byte[] HookBytes = new byte[12];

    /// <summary>
    /// Hook之前的内存数据
    /// </summary>
    public byte[] OriginBytes = new byte[12];

    /// <summary>
    /// 对话框句柄
    /// </summary>
    private IntPtr pFunc;

    /// <summary>
    /// Hook标志
    /// </summary>
    private bool Flag;

    /// <summary>
    /// 恢复密码
    /// </summary>
    private void RecoverBytes()
    {
        try
        {
            if (Flag)
            {
                Marshal.Copy(OriginBytes, 0, pFunc, 12);
            }
        }
        catch (Exception)
        {
            return;
        }
    }

    /// <summary>
    /// 定义对话框委托
    /// </summary>
    /// <param name="hInstance"></param>
    /// <param name="pTemplateName"></param>
    /// <param name="hWndParent"></param>
    /// <param name="lpDialogFunc"></param>
    /// <param name="dwInitParam"></param>
    /// <returns></returns>
    private delegate int DelegateDialogBoxParam(IntPtr hInstance, IntPtr pTemplateName, IntPtr hWndParent, IntPtr lpDialogFunc, IntPtr dwInitParam);

    /// <summary>
    /// 对话框委托(注意必须使用静态方法)
    /// </summary>
    private static DelegateDialogBoxParam MsgBoxDelegate;

    /// <summary>
    /// VBE密码Hook函数
    /// </summary>
    /// <returns></returns>
    private bool VbePasswordHook()
    {
        try
        {
            byte[] TmpBytes = new byte[12]; // 临时Hook数据
            IntPtr p;
            byte osi = 1; // 对话框句柄
            int OriginProtect = 0; // 原虚拟保护句柄
            bool result = false; // 给函数默认标志
            pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA"); // 获取导出函数的内存地址
            if (VirtualProtect(pFunc, 12, 0x40, ref OriginProtect) != IntPtr.Zero) // 标准api hook过程之一: 修改内存属性,使其可写
            {
                Marshal.Copy(pFunc, TmpBytes, 0, osi + 1); // 把内存数据复制到临时Hook变量
                if (TmpBytes[osi] != 0xB8) // 判断是否已经hook,看看API的第一个字节是否为&H68,
                {
                    Marshal.Copy(pFunc, OriginBytes, 0, 12); // 备份原始Hook数据
                    // ---------通过委托获取对话框数据
                    MsgBoxDelegate = new DelegateDialogBoxParam(MyDialogBoxParam);
                    p = Marshal.GetFunctionPointerForDelegate(MsgBoxDelegate);
                    // --------设置Hook数据到变量数组0维度
                    if (osi != 0) HookBytes[0] = 0x48;
                    HookBytes[osi] = 0xB8;
                    osi++;
                    // ---------
                    // 设置Hook数据到变量数组1-4维度
                    byte[] bytes = BitConverter.GetBytes(p.ToInt64());
                    Array.Copy(bytes, 0, HookBytes, osi, 4 * osi);
                    // -------
                    HookBytes[osi + 4 * osi] = 0xFF;
                    HookBytes[osi + 4 * osi + 1] = 0xE0;
                    // ---------
                    Marshal.Copy(HookBytes, 0, pFunc, 12); // 把Hook后的数据复制到对话框内存
                    // ------
                    Flag = true;
                    result = true;
                }
            }
            return result;
        }
        catch (Exception ex)
        {
            throw new Exception("解除密码时出现异常:" + ex.Message);
        }
    }

    /// <summary>
    /// 自定义对话框返回结果
    /// </summary>
    /// <param name="hInstance"></param>
    /// <param name="pTemplateName"></param>
    /// <param name="hWndParent"></param>
    /// <param name="lpDialogFunc"></param>
    /// <param name="dwInitParam"></param>
    /// <returns></returns>
    private int MyDialogBoxParam(IntPtr hInstance, IntPtr pTemplateName, IntPtr hWndParent, IntPtr lpDialogFunc, IntPtr dwInitParam)
    {
        try
        {
            if (pTemplateName.ToInt32() == 4070) // 有程序调用DialogBoxParamA装入4070号对话框,这里我们直接返回1,让VBE以为密码正确了
            {
                return 1;
            }
            else // 有程序调用DialogBoxParamA,但装入的不是4070号对话框,这里我们调用RecoverBytes函数恢复原来函数的功能,在进行原来的函数
            {
                RecoverBytes();
                int result = DialogBoxParam(hInstance, pTemplateName, hWndParent, lpDialogFunc, dwInitParam);
                VbePasswordHook(); // 原来的函数执行完毕,再次hook
                return result;
            }
        }
        catch (Exception ex)
        {
            throw new Exception("解除密码时出现异常:" + ex.Message);
        }
    }

    /// <summary>
    /// 解除VBA密码
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