using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using WpfHelpers;
using static WpfHelpers.ViewModelBase;

namespace BaseApp
{
    public class Validator : ViewModelBase
    {

    }
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public App()
        {
            this.InitializeComponent();
            log4net.Config.XmlConfigurator.Configure();
            StartupUri = new Uri("/BaseApp;component/MainWindow.xaml", UriKind.Relative);


        }
        //<!--StartupUri="MainWindow.xaml"-->
        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            log.Error(string.Format("Exception:{0}, StackTrace:{1}, Message:{2}", e.Exception.ToString(), e.Exception.StackTrace, e.Exception.Message));
            // Signal that we handled things--prevents Application from exiting
            e.Handled = true;
        }
    }
}
