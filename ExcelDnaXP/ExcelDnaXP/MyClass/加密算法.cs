using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDnaXP.MyClass
{
    public static class 加密算法
    {
        private const string RegKeyName = "ActivationCode";// 注册码的键名

        private static string GenerateMachineCode()
        {
            var sha = SHA256.Create();
            var rawCode = $"{Environment.MachineName}{Environment.UserName}";
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(rawCode));
            return BitConverter.ToString(hash).Replace("-", "").Substring(0, 16);
        }

        public static int ProcessRegistration(string inputCode)
        {
            try
            {
                int result = 1;
                var machineCode = GenerateMachineCode();

                if (inputCode == machineCode)
                {
                    SaveRegistration(machineCode);
                    //_isRegistered = true;
                    //Ribbon.Invalidate();
                    //MessageBox.Show("注册成功！");
                    return result;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        private static string ProtectString(string plainText)
        {
            byte[] plainBytes = Encoding.UTF8.GetBytes(plainText);
            byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.LocalMachine);
            return Convert.ToBase64String(encryptedBytes);
        }

        private static void SaveRegistration(string code)
        {
            try
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var encrypted = ProtectString(code);

                if (config.AppSettings.Settings[RegKeyName] == null)
                {
                    config.AppSettings.Settings.Add(RegKeyName, encrypted);
                }
                else
                {
                    config.AppSettings.Settings[RegKeyName].Value = encrypted;
                }

                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (ConfigurationErrorsException ex)
            {
                throw new Exception($"配置文件保存失败，请检查写入权限。详细信息: {ex.Filename}");
            }
        }
    }
}