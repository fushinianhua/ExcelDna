using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;

namespace ExcelDemo.MyCalss
{
    /// <summary>
    /// 管理应用程序配置的辅助类（包括 appsettings、connectionStrings 和自定义部分）。
    /// </summary>
    internal static class ConfigHelper
    {
        private static Configuration configuration; // 静态配置对象

        /// <summary>
        /// 静态构造函数 - 初始化配置文件路径。
        /// </summary>
        static ConfigHelper()
        {
            //execonfig.ExeConfigFilename = @"APP.config"; // 设置配置文件的相对路径
            configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        }

        /// <summary>
        /// 获取 connectionStrings 配置节中的连接字符串。
        /// </summary>
        /// <param name="name">连接字符串名称</param>
        /// <returns>连接字符串</returns>
        ///

        /// <summary>
        /// 获取自定义配置节的值。
        /// </summary>
        /// <param name="sectionName">配置节名称</param>
        /// <returns>配置节的值</returns>
        public static object GetCustomSection(string sectionName)
        {
            return configuration.GetSection(sectionName);
        }

        /// <summary>
        /// 获取 connectionStrings 配置节中的连接字符串。
        /// </summary>
        /// <param name="name">连接字符串名称</param>
        /// <returns>连接字符串</returns>
        public static string GetConnectionString(string name)
        {
            return configuration.ConnectionStrings.ConnectionStrings[name]?.ConnectionString;
        }

        /// <summary>
        /// AppSettings的
        /// </summary>
        public static class Appsettings
        { /// <summary>
          /// AppSettings键的数量
          /// </summary>
            public static int KeyCount = 0;

            private static readonly KeyValueConfigurationCollection KeyValues;

            /// <summary>
            ///APPsettings的节点
            /// </summary>
            private static AppSettingsSection appSettingsSection;

            static Appsettings()
            {
                appSettingsSection = configuration.AppSettings;
                KeyCount = appSettingsSection.Settings.Count;
                KeyValues = appSettingsSection.Settings;
            }

            /// <summary>
            /// 增加新的键值对可以直接向 NameValueCollection 中添加新项，并确保将其保存到配置文件中
            /// </summary>
            /// <param name="key">增加的键名</param>
            /// <param name="value">增加的值</param>
            public static void AddAppSetting(string key, string value)
            {
                if (KeyValues[key] == null)
                {
                    configuration.AppSettings.Settings.Add(key, value);
                    configuration.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    KeyCount += 1;
                }
            }

            /// <summary>
            /// 获取AppSettings的所有键
            /// </summary>
            /// <returns></returns>
            public static string[] GetAllkeys()
            {
                try
                {
                    if (KeyCount == 0)
                    {
                        return new string[] { };
                    }
                    string[] str = new string[KeyCount];
                    int a = 0;
                    foreach (KeyValueConfigurationElement item in KeyValues)
                    {
                        str[a] = item.Key.ToString();
                        a += 1;
                    }
                    return str;
                }
                catch
                {
                    return new string[] { };
                }
            }

            /// <summary>
            /// 获取AppSettings的所有值
            /// </summary>
            /// <returns>返回一个string数组</returns>
            public static string[] GetAllValues()
            {
                try
                {
                    if (KeyCount == 0)
                    {
                        return new string[] { };
                    }
                    List<string> liststr = new List<string>() { };
                    foreach (KeyValueConfigurationElement item in KeyValues)
                    {
                        liststr.Add(item.Value.ToString());
                    }
                    return liststr.ToArray();
                }
                catch
                {
                    return new string[] { };
                }
            }

            /// <summary>
            /// 获取键值对的值
            /// </summary>
            /// <param name="key">需要查找的键名</param>
            /// <returns>返回键名对应的值</returns>
            public static string GetValue(string key)
            {
                string Result = "";
                try
                {
                    if (KeyCount == 0)
                    {
                        Result = "";
                    }
                    if (KeyValues[key] != null)
                    {
                        Result = configuration.AppSettings.Settings[key].Value;
                    }

                    return Result;
                }
                catch
                {
                    return "";
                }
            }

            /// <summary>
            /// 移出删除某个键值对
            /// </summary>
            /// <param name="key">需要移出的键名</param>
            public static void RemoveAppSetting(string key)
            {
                if (configuration.AppSettings.Settings[key] != null)
                {
                    configuration.AppSettings.Settings.Remove(key);
                    configuration.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                    KeyCount -= 1;
                }
            }

            /// <summary>
            /// 修改现有的键值对
            /// </summary>
            /// <param name="key">要修改的键名</param>
            /// <param name="value">要修改的值</param>
            public static void UpdateAppSetting(string key, string value)
            {
                Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                if (configuration.AppSettings.Settings[key] != null)
                {
                    configuration.AppSettings.Settings[key].Value = value;
                    configuration.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                }
            }
        }

        public static class connectionString
        {
            private static ConnectionStringsSection ConnectionStringsSection; // 连接字符串配置节对象
            private static ConnectionStringSettingsCollection connectionStringSettings;

            static connectionString()
            {
                ConnectionStringsSection = configuration.ConnectionStrings; // 获取连接字符串配置节
                connectionStringSettings = ConnectionStringsSection.ConnectionStrings; // 获取连接字符串集合
            }

            /// <summary>
            /// 添加连接字符串到配置文件中。
            /// </summary>
            /// <param name="name">连接字符串的名称</param>
            /// <param name="value">连接字符串的值</param>
            /// <param name="providername">提供程序名称（可选）</param>
            public static void Add(string name, string value, string providername)
            {
                if (connectionStringSettings[name] == null)
                {
                    ConnectionStringSettings connectionStringsSection = new ConnectionStringSettings(name, value, providername);
                    connectionStringSettings.Add(connectionStringsSection);
                    ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("connectionStrings");
                }
            }

            /// <summary>
            /// 更新现有连接字符串的值。
            /// </summary>
            /// <param name="name">需要更新的连接字符串的名称</param>
            /// <param name="value">新的连接字符串值</param>
            public static void Update(string name, string value)
            {
                if (connectionStringSettings[name] != null)
                {
                    connectionStringSettings[name].ConnectionString = value;
                    ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("connectionStrings");
                }
            }

            /// <summary>
            /// 获取配置文件中特定连接字符串的值。
            /// </summary>
            /// <param name="name">需要获取的连接字符串的名称</param>
            /// <returns>连接字符串的值</returns>
            public static string GetConnectionString(string name)
            {
                return ConfigurationManager.ConnectionStrings[name]?.ConnectionString;
            }

            /// <summary>
            /// 设置或添加配置文件中的连接字符串。
            /// </summary>
            /// <param name="name">连接字符串的名称</param>
            /// <param name="connectionString">连接字符串的值</param>
            public static void SetConnectionString(string name, string connectionString)
            {
                if (connectionStringSettings[name] != null)
                {
                    connectionStringSettings[name].ConnectionString = connectionString;
                }
                else
                {
                    var cs = new ConnectionStringSettings(name, connectionString);
                    connectionStringSettings.Add(cs);
                }
                ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("connectionStrings");
            }
        }
    }
}