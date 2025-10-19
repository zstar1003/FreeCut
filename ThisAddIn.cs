using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Extensibility;

namespace FreeCut
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789ABC")]
    [ProgId("FreeCut.ThisAddIn")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ThisAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        // 完全空的实现 - 用于测试COM加载
        private static object powerPointApp;

        public static object PowerPointApplication => powerPointApp;

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            powerPointApp = application;

            // 尝试写日志
            try
            {
                string logPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "FreeCut_load.log");
                System.IO.File.AppendAllText(logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] OnConnection called - TEST VERSION\n");
            }
            catch { }
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                string logPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "FreeCut_load.log");
                System.IO.File.AppendAllText(logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] GetCustomUI called - TEST VERSION\n");
            }
            catch { }

            return string.Empty;
        }

        // 添加必要的静态方法以满足编译
        public static List<dynamic> GetSelectedSlides()
        {
            return new List<dynamic>();
        }

        public static List<dynamic> GetAllSlides()
        {
            return new List<dynamic>();
        }

        public static string ExportSlideToImage(dynamic slide, string outputPath, int dpi = 300)
        {
            return string.Empty;
        }
    }
}

