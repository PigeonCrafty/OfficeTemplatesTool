using System;
using System.Windows;
using System.Windows.Threading;

namespace WPF.Themes.Demo
{
    /// <summary>
    ///     Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            if (e.Exception is InvalidOperationException)
                e.Handled = true;
        }
    }
}