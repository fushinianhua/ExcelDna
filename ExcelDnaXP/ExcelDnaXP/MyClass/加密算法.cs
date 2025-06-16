using Radiant.Properties;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace Radiant
{
    public static class 加密算法
    {
        // 静态变量，存储硬件信息和加密密钥
        public static string CPUID = "";

        public static string key = "";
        public static string 激活码 = "";
        public static string 注册码 = "";
        public static string 密钥 = "21218308";
        private static bool _isInitialized = false;

        // 存储生成的机器码，确保在初始化后保持不变
        private static string _generatedMachineCode = "";

        /// <summary>
        /// 初始化加密算法，确保在使用前正确加载配置和硬件信息
        /// </summary>
        public static void Initialize()
        {
            if (_isInitialized) return;

            try
            {
                // 从配置或其他安全位置加载密钥
                LoadEncryptionKey();

                // 获取硬件信息
                CPUID = 获取CPUID();

                // 生成机器码 - 确保只生成一次
                if (string.IsNullOrEmpty(_generatedMachineCode))
                {
                    _generatedMachineCode = 生成机器码(CPUID);
                }
                注册码 = _generatedMachineCode;

                _isInitialized = true;
            }
            catch (Exception ex)
            {
                LogError(ex, "加密算法初始化失败");
                throw new Exception("加密算法初始化异常，请联系技术支持", ex);
            }
        }

        /// <summary>
        /// 安全加载加密密钥，优先从配置文件，失败时尝试从注册表或其他位置
        /// </summary>
        private static void LoadEncryptionKey()
        {
            try
            {
                // 尝试从配置文件加载
                key = Settings.Default.密钥;

                if (string.IsNullOrEmpty(key))
                {
                    // 从注册表加载
                    key = LoadKeyFromRegistry();

                    if (string.IsNullOrEmpty(key))
                    {
                        // 生成临时密钥（仅用于演示，实际生产环境不建议）
                        key = GenerateTemporaryKey();
                        LogError(null, "使用临时加密密钥，建议配置正式密钥");
                    }
                    else
                    {
                        // 保存到配置以便下次使用
                        Settings.Default.密钥 = key;
                        Settings.Default.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录错误但继续执行，使用临时密钥
                LogError(ex, "加载加密密钥失败，使用临时密钥");
                key = GenerateTemporaryKey();
            }
        }

        /// <summary>
        /// 从注册表加载加密密钥
        /// </summary>
        private static string LoadKeyFromRegistry()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key =
                    Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\YourCompany\\YourAddin"))
                {
                    if (key != null)
                    {
                        return key.GetValue("EncryptionKey") as string;
                    }
                }

                // 尝试从本地机器注册表读取（需要管理员权限）
                using (Microsoft.Win32.RegistryKey key =
                    Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\YourCompany\\YourAddin"))
                {
                    if (key != null)
                    {
                        return key.GetValue("EncryptionKey") as string;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogError(ex, "从注册表加载密钥失败");
                return null;
            }
        }

        /// <summary>
        /// 生成临时加密密钥（仅用于调试和演示）
        /// </summary>
        private static string GenerateTemporaryKey()
        {
            // 注意：此方法仅用于演示，实际生产环境应使用安全的密钥管理
            using (RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider())
            {
                byte[] keyBytes = new byte[16];
                rng.GetBytes(keyBytes);
                return Convert.ToBase64String(keyBytes);
            }
        }

        /// <summary>
        /// 获取CPUID或其他硬件标识
        /// </summary>
        public static string 获取CPUID()
        {
            try
            {
                // 首先尝试使用WMI获取CPU信息
                string cpuInfo = GetProperties(new string[] { "Win32_Processor", "Name", "Manufacturer", "ProcessorId" });

                if (!string.IsNullOrEmpty(cpuInfo))
                {
                    return cpuInfo;
                }

                // 如果WMI失败，尝试使用替代方法
                LogError(null, "WMI获取CPU信息失败，尝试替代方法");
                return GetAlternativeHardwareId();
            }
            catch (ManagementException me)
            {
                LogError(me, "WMI访问异常，尝试替代方法");
                return GetAlternativeHardwareId();
            }
            catch (Exception ex)
            {
                LogError(ex, "获取硬件标识异常，使用随机ID");
                // 返回一个基于时间戳的唯一ID，确保每次运行不同
                return "TEMP-" + DateTime.Now.Ticks.ToString("X");
            }
        }

        /// <summary>
        /// 当WMI无法访问时，使用替代方法获取硬件标识
        /// </summary>
        private static string GetAlternativeHardwareId()
        {
            try
            {
                // 尝试获取主板序列号
                string motherboardId = GetProperties(new string[] { "Win32_BaseBoard", "SerialNumber" });
                if (!string.IsNullOrEmpty(motherboardId))
                {
                    return motherboardId;
                }

                // 尝试获取磁盘序列号
                string diskId = GetProperties(new string[] { "Win32_DiskDrive", "Model", "SerialNumber" });
                if (!string.IsNullOrEmpty(diskId))
                {
                    return diskId;
                }

                // 如果都失败，使用网卡MAC地址
                string macAddress = GetMacAddress();
                if (!string.IsNullOrEmpty(macAddress))
                {
                    return macAddress;
                }

                // 如果所有方法都失败，生成一个基于机器名和当前时间的唯一ID
                return Environment.MachineName + "-" + DateTime.Now.Ticks.ToString("X");
            }
            catch (Exception ex)
            {
                LogError(ex, "所有硬件标识方法失败，使用随机ID");
                return Guid.NewGuid().ToString();
            }
        }

        /// <summary>
        /// 获取网卡MAC地址
        /// </summary>
        private static string GetMacAddress()
        {
            try
            {
                string macAddresses = "";

                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"))
                {
                    ManagementObjectCollection moCollection = searcher.Get();

                    foreach (ManagementObject mo in moCollection.Cast<ManagementObject>())
                    {
                        macAddresses += mo["MacAddress"].ToString() + ",";
                        mo.Dispose();
                    }
                }

                if (!string.IsNullOrEmpty(macAddresses))
                {
                    return macAddresses.TrimEnd(',');
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取WMI数据
        /// </summary>
        private static string GetProperties(string[] wmiData)
        {
            try
            {
                StringBuilder properties = new StringBuilder();
                string query = GenerateQuery(wmiData);

                using (ManagementObjectSearcher moSearcher = new ManagementObjectSearcher("root\\CIMV2", query))
                {
                    using (ManagementObjectCollection moCollection = moSearcher.Get())
                    {
                        foreach (ManagementObject mbObject in moCollection)
                        {
                            using (mbObject)
                            {
                                for (int i = 1; i < wmiData.Length; i++)
                                {
                                    if (mbObject[wmiData[i]] != null)
                                    {
                                        properties.Append(",");
                                        properties.Append(mbObject[wmiData[i]].ToString().Trim());
                                    }
                                }
                            }
                            properties.Append(";");
                        }
                    }
                }

                if (properties.Length > 1)
                {
                    return properties.ToString().Substring(1);
                }

                return null;
            }
            catch (Exception ex)
            {
                LogError(ex, "获取WMI属性失败");
                return null;
            }
        }

        /// <summary>
        /// 生成WMI查询语句
        /// </summary>
        private static string GenerateQuery(string[] wmiData)
        {
            StringBuilder query = new StringBuilder();
            string wmiClass = string.Empty;

            query.Append("SELECT ");

            for (int i = 0; i < wmiData.Length; i++)
            {
                if (i == 0)
                {
                    wmiClass = wmiData[i];
                }
                else
                {
                    query.Append(i < wmiData.Length - 1 ? $"{wmiData[i]}, " : $"{wmiData[i]} ");
                }
            }

            query.Append($"FROM {wmiClass}");
            return query.ToString();
        }

        /// <summary>
        /// 生成机器码
        /// </summary>
        public static string 生成机器码(string input)
        {
            try
            {
                if (string.IsNullOrEmpty(input))
                {
                    throw new ArgumentNullException(nameof(input), "生成机器码的输入不能为空");
                }

                using (MD5 md5 = MD5.Create())
                {
                    // 将输入字符串和密钥组合
                    byte[] inputBytes = Encoding.UTF8.GetBytes(input + key);

                    // 计算哈希值
                    byte[] hashBytes = md5.ComputeHash(inputBytes);

                    // 转换为十六进制字符串
                    StringBuilder hexStringBuilder = new StringBuilder();
                    foreach (byte b in hashBytes)
                    {
                        hexStringBuilder.Append(b.ToString("x2"));
                    }

                    string hexString = hexStringBuilder.ToString();

                    // 格式化机器码，每4位用短横线分隔
                    StringBuilder formattedStringBuilder = new StringBuilder();
                    for (int i = 0; i < hexString.Length; i += 4)
                    {
                        if (i > 0)
                        {
                            formattedStringBuilder.Append("-");
                        }

                        int length = Math.Min(4, hexString.Length - i);
                        formattedStringBuilder.Append(hexString.Substring(i, length));
                    }

                    注册码 = formattedStringBuilder.ToString().ToUpper();
                    return 注册码;
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "生成机器码失败");
                throw new Exception("生成机器码异常", ex);
            }
        }

        /// <summary>
        /// 对字符串进行加密并格式化结果
        /// </summary>
        public static string EncryptAndFormat(string input, string separator = "-", int segmentLength = 4)
        {
            // 输入验证
            if (string.IsNullOrEmpty(input))
            {
                throw new ArgumentException("输入字符串不能为空。", nameof(input));
            }

            if (string.IsNullOrEmpty(separator))
            {
                separator = "-";
            }

            if (segmentLength <= 0)
            {
                segmentLength = 4;
            }

            try
            {
                using (MD5 md5 = MD5.Create())
                {
                    // 计算哈希值
                    byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                    byte[] hashBytes = md5.ComputeHash(inputBytes);

                    // 转换为十六进制字符串
                    StringBuilder hexStringBuilder = new StringBuilder();
                    foreach (byte b in hashBytes)
                    {
                        hexStringBuilder.Append(b.ToString("x2"));
                    }

                    string hexString = hexStringBuilder.ToString();

                    // 格式化字符串
                    StringBuilder formattedStringBuilder = new StringBuilder();
                    for (int i = 0; i < hexString.Length; i += segmentLength)
                    {
                        if (i > 0)
                        {
                            formattedStringBuilder.Append(separator);
                        }

                        int length = Math.Min(segmentLength, hexString.Length - i);
                        formattedStringBuilder.Append(hexString.Substring(i, length));
                    }

                    激活码 = formattedStringBuilder.ToString();
                    return 激活码;
                }
            }
            catch (CryptographicException ex)
            {
                LogError(ex, "加密过程中发生加密错误");
                return null;
            }
            catch (Exception ex)
            {
                LogError(ex, "发生未知错误");
                return null;
            }
        }

        /// <summary>
        /// 验证注册码是否有效
        /// </summary>
        public static bool ValidateRegistration(string userEnteredCode)
        {
            try
            {
                // 确保算法已初始化
                if (!_isInitialized)
                {
                    Initialize();
                }

                // 生成基于当前机器码的激活码（关键修正点）
                string machineCode = 注册码;
                string expectedCode = EncryptAndFormat(machineCode);

                // 验证逻辑：用户输入的注册码与预期激活码匹配，或等于备用密钥
                return !string.IsNullOrEmpty(expectedCode) &&
                       (expectedCode.Equals(userEnteredCode, StringComparison.OrdinalIgnoreCase) ||
                        userEnteredCode.Equals(密钥, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                LogError(ex, "注册码验证失败");
                return false;
            }
        }

        /// <summary>
        /// 记录错误日志
        /// </summary>
        private static void LogError(Exception ex, string message)
        {
            try
            {
                // 构建日志消息
                StringBuilder logMessage = new StringBuilder();
                logMessage.AppendLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                logMessage.AppendLine($"位置: 加密算法 - {message}");

                if (ex != null)
                {
                    logMessage.AppendLine($"异常类型: {ex.GetType().FullName}");
                    logMessage.AppendLine($"异常消息: {ex.Message}");

                    if (!string.IsNullOrEmpty(ex.StackTrace))
                    {
                        logMessage.AppendLine($"堆栈跟踪: {ex.StackTrace}");
                    }

                    if (ex.InnerException != null)
                    {
                        logMessage.AppendLine($"内部异常: {ex.InnerException.Message}");
                    }
                }

                logMessage.AppendLine(new string('-', 80));

                // 写入日志文件
                string logFilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "YourAddin",
                    "EncryptionLog.txt");

                // 确保目录存在
                string logDirectory = Path.GetDirectoryName(logFilePath);
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // 追加日志
                File.AppendAllText(logFilePath, logMessage.ToString());
            }
            catch
            {
                // 日志记录失败，无法做更多操作
            }
        }
    }
}