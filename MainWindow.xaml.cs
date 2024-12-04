using System.Diagnostics;
using System.Windows;

namespace BaseApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Process[] pname = Process.GetProcessesByName(assemblyName);
            if (pname.Length > 1)
            {
                MessageBox.Show("Existing instance already running...");
                System.Windows.Application.Current.Shutdown();
            }
            ViewModels.Bootstrap dataContext = new ViewModels.Bootstrap();
            this.DataContext = dataContext;
        }

        private void btnMinimize_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.WindowState = System.Windows.WindowState.Minimized;
        }
    }
}
