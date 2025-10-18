using System;
using System.Windows;
using System.Windows.Threading;

namespace FreeCut
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 设置全局异常处理
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            // 设置应用程序信息
            this.ShutdownMode = ShutdownMode.OnMainWindowClose;
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(
                $"应用程序发生错误：\n\n{e.Exception.Message}\n\n详细信息：\n{e.Exception.StackTrace}",
                "FreeCut - 错误",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            MessageBox.Show(
                $"应用程序发生严重错误：\n\n{exception?.Message ?? "未知错误"}",
                "FreeCut - 严重错误",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }
}