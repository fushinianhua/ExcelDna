using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Radiant.Myform
{
    public partial class 网页 : Form
    {
        public string url { get; set; }
        private WebView2 webView2 = null;
        private bool isWebViewInitialized = false;

        public 网页()
        {
            InitializeComponent();
        }

        private void 网页_Load(object sender, EventArgs e)
        {
            try
            {
                webView2 = new WebView2();
                webView2.Dock = DockStyle.Fill;
                this.panel1.Controls.Add(webView2);

                webView2.NavigationCompleted += WebView2_NavigationCompleted;
                webView2.CoreWebView2InitializationCompleted += WebView2_CoreWebView2InitializationCompleted;

                InitializeWebView2Async();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化网页控件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void InitializeWebView2Async()
        {
            try
            {
                if (!await CheckWebView2Runtime())
                {
                    ShowRuntimeMissingMessage();
                    return;
                }

                // 设置数据目录
                string dataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "msedge.exe",  // 替换为你的应用名称
                    "WebView2Data");

                // 创建目录
                Directory.CreateDirectory(dataFolder);

                // 记录数据目录路径（用于调试）
                Console.WriteLine($"WebView2 数据目录: {dataFolder}");

                // 检查磁盘空间
                long availableSpace = GetAvailableDiskSpace(dataFolder);
                if (availableSpace < 100 * 1024 * 1024)  // 小于100MB
                {
                    MessageBox.Show($"警告: 磁盘空间不足 ({availableSpace / (1024 * 1024):N0} MB)",
                        "资源警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // 创建带数据目录的环境
                var env = await CoreWebView2Environment.CreateAsync(null, dataFolder);
                await webView2.EnsureCoreWebView2Async(env);

                webView2.CoreWebView2.Navigate(url ?? "https://www.baidu.com");
                isWebViewInitialized = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WebView2 初始化异常: {ex}");
                MessageBox.Show($"WebView2 初始化失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CleanupWebView();
            }
        }

        private void ShowRuntimeMissingMessage()
        {
            MessageBox.Show(
                "未检测到 WebView2 运行时。\n" +
                "请访问 https://developer.microsoft.com/en-us/microsoft-edge/webview2/ 下载并安装。",
                "缺少依赖",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        // 兼容旧版本 C# 的运行时检查方法
        private async Task<bool> CheckWebView2Runtime()
        {
            try
            {
                // 尝试创建环境
                await CoreWebView2Environment.CreateAsync(null, null);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"检查 WebView2 运行时失败: {ex.Message}");

                // 备选：检查注册表（兼容 C# 7.3 写法）
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    try
                    {
                        // 传统 Using 语句
                        RegistryKey key = null;
                        try
                        {
                            key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                                @"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}");

                            object value = key?.GetValue("pv");
                            return key != null && value != null;
                        }
                        finally
                        {
                            // 确保释放资源
                            if (key != null)
                            {
                                key.Close();
                            }
                        }
                    }
                    catch { }
                }

                return false;
            }
        }

        private long GetAvailableDiskSpace(string path)
        {
            try
            {
                DriveInfo drive = new DriveInfo(Path.GetPathRoot(path));
                return drive.AvailableFreeSpace;
            }
            catch
            {
                return -1;
            }
        }

        private void WebView2_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            try
            {
                if (e.IsSuccess)
                {
                    this.Text = $"网页 - {webView2.CoreWebView2.DocumentTitle}";
                }
                else
                {
                    MessageBox.Show($"网页加载失败: {e.WebErrorStatus}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导航事件处理失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WebView2_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {
                if (!e.IsSuccess)
                {
                    MessageBox.Show($"WebView2 初始化失败: {e.InitializationException.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化事件处理失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CleanupWebView()
        {
            try
            {
                if (webView2 != null)
                {
                    if (webView2.CoreWebView2 != null)
                    {
                        webView2.CoreWebView2.Stop();
                    }

                    if (this.panel1.Controls.Contains(webView2))
                    {
                        this.panel1.Controls.Remove(webView2);
                    }

                    webView2.Dispose();
                    webView2 = null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理 WebView2 失败: {ex.Message}");
            }
        }

        private void 网页_FormClosing(object sender, FormClosingEventArgs e)
        {
            CleanupWebView();
        }
    }
}