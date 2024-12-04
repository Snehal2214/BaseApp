using System;
using System.IO;
using System.Windows.Input;
using WpfHelpers;
using WpfHelpers.Controls;

namespace BaseApp.ViewModels
{
    public partial class Bootstrap : ViewModelBase
    {
        readonly log4net.ILog SpecificLogger = log4net.LogManager.GetLogger("Summary");

        public static string settingPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Settings.xlsx");

        public static int UserRole { get; set; }
        public static string UserName { get; set; }
        bool _LoginVisible = true;
        public bool LoginVisible
        {
            get
            {
                return _LoginVisible;
            }
            set
            {
                _LoginVisible = value;
                OnPropertyChanged("LoginVisible");
            }
        }
        bool reportvisible = false;
        public bool ReportVisible
        {
            get
            {
                return reportvisible;
            }
            set
            {
                reportvisible = value;
                OnPropertyChanged("ReportVisible");
            }
        }
        bool _LogoutVisible = false;
        public bool LogoutVisible
        {
            get
            {
                return _LogoutVisible;
            }
            set
            {
                _LogoutVisible = value;
                OnPropertyChanged("LogoutVisible");
            }
        }
        bool menusettingsvisible = false;
        public bool SettingsVisible { get { return menusettingsvisible; } set { menusettingsvisible = value; OnPropertyChanged("SettingsVisible"); } }
        bool dashboardVisible = false;
        public bool DashboardVisible
        {
            get { return dashboardVisible; }
            set
            {
                dashboardVisible = value;
                OnPropertyChanged("DashboardVisible");
            }
        }
        object dashBoardContext = null;
        object reportContext = null;
        object _Context;
        public object Context
        {
            get { return _Context; }
            set
            {
                _Context = value;
                OnPropertyChanged("Context");
            }
        }
    }
    public partial class Bootstrap : ViewModelBase
    {
        public ICommand LogoutUser { get; set; }
        public ICommand LoginUser { get; set; }
        public ICommand ShowView { get; set; }
        ViewModels.Login userlLoginVM = new Login();
        public Bootstrap()
        {
            InitializeCommands();
            userlLoginVM.LoginEvent += userlLoginVM_LoginEvent;
            Context = userlLoginVM;
            //LoginVisible = false;
            //LogoutVisible = true;
            //ReportVisible = SettingsVisible = DashboardVisible = true;
            //dashBoardContext = dashBoardContext != null ? dashBoardContext : new ViewModels.Dashboard();
            //Context = dashBoardContext;
        }
        void InitializeCommands()
        {
            LogoutUser = new DelegateCommand((buttonTag) =>
            {
                SpecificLogger.Debug($"User Logged out {UserName}\r\n");
                UserName = null;
                Context = null;
                LoginVisible = true;
                LogoutVisible = SettingsVisible = DashboardVisible = ReportVisible = false;
                userlLoginVM = new Login();
                userlLoginVM.LoginEvent += userlLoginVM_LoginEvent;
                Context = userlLoginVM;
            });
            ShowView = new DelegateCommand((buttonTag) =>
            {
                if (buttonTag == null)
                {
                    return;
                }
                if (buttonTag.Equals("BaseApp.ViewModels.Dashboard"))
                {
                    // Once the dashboard viewmodel is initialized then we dont want the viewmodel to be re-initialized
                    // so resue the dashBoardContext
                    SpecificLogger.Debug($"User navigated to {buttonTag.ToString().Replace("BaseApp.ViewModels.", "")}");
                    dashBoardContext = dashBoardContext != null ? dashBoardContext : new ViewModels.Dashboard();
                    Context = dashBoardContext;
                }
                else
                {
                    SpecificLogger.Debug($"User navigated to {buttonTag.ToString().Replace("BaseApp.ViewModels.", "")}");
                    Type type = Type.GetType(buttonTag.ToString());
                    Context = Activator.CreateInstance(type);
                }
            });
        }
        void userlLoginVM_LoginEvent(object sender, Login.LoginEventArgs e)
        {
            if (e.IsValid)
            {
                UserName = e.UserName;
                SpecificLogger.Debug($"User Logged in: {e.UserName}");
                LoginVisible = false;
                LogoutVisible = true;
                if (e.Role.ToUpper() == "1")
                {
                    UserRole = 1;
                    DashboardVisible = true; SettingsVisible = true;
                    {
                        SpecificLogger.Debug($"User navigated to Dashboard");
                        dashBoardContext = new ViewModels.Dashboard();
                        Context = dashBoardContext;
                    }
                }
                else if (e.Role.ToUpper() == "2")
                {
                    UserRole = 2;
                    DashboardVisible = true; SettingsVisible = false;
                    {
                        SpecificLogger.Debug($"User navigated to Dashboard");
                        dashBoardContext = new ViewModels.Dashboard();
                        Context = dashBoardContext;
                    }
                }
                else
                {
                    SettingsVisible = DashboardVisible = ReportVisible = false;
                }
            }
        }
    }
    public class RelayCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;
        private Action DoWork;

        public RelayCommand(Action Work)
        {
            DoWork = Work;
        }


        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            DoWork();
        }
    }
}
