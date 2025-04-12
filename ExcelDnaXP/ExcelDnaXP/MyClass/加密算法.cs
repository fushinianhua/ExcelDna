using Radiant.Properties;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelDna.Integration.XlCall;

namespace Radiant
{
    public static class 加密算法
    {
        public static string CPUID = "";
        public static string key = Settings.Default.密钥;
        public static string 激活码 = ""; //CPUID
        public static string 机器码 = "";

        /// <summary>
        /// 获取CPUID
        /// </summary>
        /// <returns></returns>
        public static string 获取CPUID()
        {
            CPUID = GetProperties(new string[] { "Win32_Processor", "Name", "Manufacturer", "ProcessorId" });
            return CPUID;
        }

        /// <summary>
        /// 获取WMI数据
        /// </summary>
        /// <param name="wmiData"></param>
        /// <returns></returns>
        private static string GetProperties(string[] wmiData)
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
                                properties.Append(",");
                                properties.Append(mbObject[wmiData[i]].ToString().Trim());
                            }
                        }
                        properties.Append(";");
                    }
                }
            }
            return properties.ToString().Substring(1);
        }

        /// <summary>
        /// 生成查询
        /// </summary>
        /// <param name="wmiData"></param>
        /// <returns></returns>
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
        /// 对字符串进行 MD5 加密并格式化结果
        /// </summary>
        /// <param name="input"></param>
        /// <param name="separator"></param>
        /// <param name="segmentLength"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>

        public static string 生成机器码(string input)
        {
            try
            {
                using (MD5 md5 = MD5.Create())
                {
                    // 将输入字符串转换为字节数组
                    byte[] inputBytes = Encoding.UTF8.GetBytes(input + key);
                    // 计算哈希值
                    byte[] hashBytes = md5.ComputeHash(inputBytes);

                    // 将字节数组转换为十六进制字符串
                    StringBuilder hexStringBuilder = new StringBuilder();
                    foreach (byte b in hashBytes)
                    {
                        hexStringBuilder.Append(b.ToString("x2"));
                    }
                    string hexString = hexStringBuilder.ToString();

                    // 格式化十六进制字符串，每 segmentLength 位用 separator 分隔
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

                    机器码 = formattedStringBuilder.ToString().ToUpper();
                    return 机器码;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// 对字符串进行 MD5 加密并格式化结果
        /// </summary>
        /// <param name="input"></param>
        /// <param name="separator"></param>
        /// <param name="segmentLength"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static string EncryptAndFormat(string input, string separator = "-", int segmentLength = 4)
        {
            // 输入验证
            if (string.IsNullOrEmpty(input))
            {
                throw new ArgumentException("输入字符串不能为空。", nameof(input));
            }
            if (string.IsNullOrEmpty(separator))
            {
                throw new ArgumentException("分隔符不能为空。", nameof(separator));
            }
            if (segmentLength <= 0)
            {
                throw new ArgumentException("分段长度必须为正整数。", nameof(segmentLength));
            }

            try
            {
                using (MD5 md5 = MD5.Create())
                {
                    // 将输入字符串转换为字节数组
                    byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                    // 计算哈希值
                    byte[] hashBytes = md5.ComputeHash(inputBytes);

                    // 将字节数组转换为十六进制字符串
                    StringBuilder hexStringBuilder = new StringBuilder();
                    foreach (byte b in hashBytes)
                    {
                        hexStringBuilder.Append(b.ToString("x2"));
                    }
                    string hexString = hexStringBuilder.ToString();

                    // 格式化十六进制字符串，每 segmentLength 位用 separator 分隔
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
                Console.WriteLine($"加密过程中发生加密错误: {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生未知错误: {ex.Message}");
                return null;
            }
        }
    }
}