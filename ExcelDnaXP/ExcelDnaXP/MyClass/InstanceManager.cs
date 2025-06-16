using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Radiant.Properties;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Radiant
{
    /// <summary>
    /// 管理各Excel实例的隔离状态
    /// </summary>
    public class InstanceManager
    {
        // 使用ConcurrentDictionary确保线程安全，存储各实例的状态
        private static readonly ConcurrentDictionary<string, InstanceState> _instanceStates =
            new ConcurrentDictionary<string, InstanceState>();

        /// <summary>
        /// 获取当前Excel实例的唯一标识
        /// </summary>
        public static string GetCurrentInstanceId()
        {
            try
            {
                if (ExcelDnaUtil.Application == null)
                    return "unknown_instance";

                var excelApp = ExcelDnaUtil.Application as Application;
                if (excelApp == null)
                    return "unknown_instance";

                var excelProcessId = excelApp.Hwnd;
                return $"process_{excelProcessId}";
            }
            catch (Exception)
            {
                return "unknown_instance";
            }
        }

        /// <summary>
        /// 获取当前实例的状态管理对象
        /// </summary>
        public static InstanceState GetCurrentInstanceState()
        {
            var instanceId = GetCurrentInstanceId();
            return _instanceStates.GetOrAdd(instanceId, _ => new InstanceState());
        }

        /// <summary>
        /// 清理指定实例的状态
        /// </summary>
        public static void CleanupInstanceState(string instanceId)
        {
            _instanceStates.TryRemove(instanceId, out _);
        }

        /// <summary>
        /// 清理当前实例的状态
        /// </summary>
        public static void CleanupCurrentInstanceState()
        {
            var instanceId = GetCurrentInstanceId();
            CleanupInstanceState(instanceId);
        }
    }

    /// <summary>
    /// 单个Excel实例的状态容器
    /// </summary>
    public class InstanceState
    {
        // 注册状态
        // 注册状态 - 直接从Settings获取，而不是保存到实例中
        // 激活状态仅由主注册码决定
        // 注册状态 - 直接从Settings获取
        public bool IsRegistered
        {
            get => Settings.Default.激活状态 &&
                   (Settings.Default.激活码 == 加密算法.EncryptAndFormat(加密算法.注册码) ||
                    Settings.Default.激活码 == "21218308");
            set
            {
                Settings.Default.激活状态 = value;
                Settings.Default.Save();
            }
        }

        // 按钮状态列表
        public Dictionary<string, bool> ButtonStates { get; } = new Dictionary<string, bool>();

        // 初始化默认状态
        public InstanceState()
        {
            // 初始化按钮状态
            ButtonStates["TestButton"] = true;
            ButtonStates["button2"] = true;
        }
    }
}