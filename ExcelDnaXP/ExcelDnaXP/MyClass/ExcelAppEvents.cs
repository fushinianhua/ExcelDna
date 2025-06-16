using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Radiant.MyClass
{
    /// <summary>
    /// Excel应用程序事件处理
    /// </summary>
    internal class ExcelAppEvents : IDisposable
    {
        private readonly Application _excelApp;
        private readonly Action<Workbook> _onWorkbookClose;
        private bool _disposed = false;

        public ExcelAppEvents(Application excelApp, Action<Workbook> onWorkbookClose)
        {
            _excelApp = excelApp;
            _onWorkbookClose = onWorkbookClose;

            // 直接订阅事件，不使用中间委托
            _excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
        }

        private void ExcelApp_WorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            // 调用用户提供的回调，忽略cancel参数
            _onWorkbookClose?.Invoke(wb);

            // 确保不取消关闭操作
            cancel = false;
        }

        // 实现IDisposable接口
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // 清理托管资源
                }

                // 清理非托管资源
                try
                {
                    if (_excelApp != null)
                    {
                        _excelApp.WorkbookBeforeClose -= ExcelApp_WorkbookBeforeClose;
                    }
                }
                catch { }

                _disposed = true;
            }
        }

        ~ExcelAppEvents()
        {
            Dispose(false);
        }
    }
}