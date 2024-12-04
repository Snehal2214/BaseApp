using BaseApp.Models;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Windows.Input;
using WpfHelpers;
using WpfHelpers.Controls;

namespace BaseApp.ViewModels
{
    public class Login : ViewModelBase
    {
        #region Event And Handlers

        public event EventHandler<LoginEventArgs> LoginEvent;
        protected virtual void OnLogin(LoginEventArgs e)
        {
            EventHandler<LoginEventArgs> handler = LoginEvent;
            if (handler != null)
            {
                handler(this, e);
            }
        }


        public class LoginEventArgs : EventArgs
        {
            public LoginEventArgs()
            {
                LoginTime = DateTime.SpecifyKind(LoginTime, DateTimeKind.Utc);
            }
            public bool IsValid { get; set; }
            public string Role { get; set; }
            public DateTime LoginTime { get; set; }
            public string UserName { get; set; }
        }
        #endregion

        #region Commands
        public ICommand LoginUser { get; set; }
        public ICommand ChangePassword { get; set; }
        public ICommand CancelChangePassword { get; set; }
        public ICommand ChangeUserPassword { get; set; }
        #endregion

        #region Properties

        string loginID, password;
        string oldpassword, newpassword, confirmpassword = string.Empty;
        public string ConfirmPassword
        {
            get
            {
                return confirmpassword;
            }
            set
            {
                confirmpassword = value;
                OnPropertyChanged("ConfirmPassword");
            }
        }
        public string NewPassword
        {
            get
            {
                return newpassword;
            }
            set
            {
                newpassword = value;
                OnPropertyChanged("NewPassword");
            }
        }
        public string OldPassword
        {
            get
            {
                return oldpassword;
            }
            set
            {
                oldpassword = value;
                OnPropertyChanged("OldPassword");
            }
        }
        bool changepasswordnow = false;
        public bool ChangePasswordNow
        {
            get
            {
                return changepasswordnow;
            }
            set
            {
                changepasswordnow = value;
                OnPropertyChanged("ChangePasswordNow");
            }
        }
        bool _InCorrectUser = false;
        public bool InCorrectUser
        {
            get
            {
                return _InCorrectUser;
            }
            set
            {
                _InCorrectUser = value;
                OnPropertyChanged("InCorrectUser");
            }
        }
        string message = null;
        public string Message
        {
            get { return message; }
            set
            {
                message = value;
                OnPropertyChanged("Message");
            }
        }
        public string LoginID
        {
            get { return loginID; }
            set
            {
                loginID = value;
                OnPropertyChanged("LoginID");
            }
        }
        public string Password { get { return password; } set { password = value; OnPropertyChanged("Password"); } }
        #endregion

        public Login()
        {
            InitializeCommands();

        }


        private void InitializeCommands()
        {
            CancelChangePassword = new DelegateCommand((param) =>
            {

                //umgr.AddLog(DateTime.Now.ToString("dd-MMM-yyyy"), "Cancel Change Password", Bootstrap.UserName);
                ChangePasswordNow = false;
                Message = "";
            });
            ChangeUserPassword = new DelegateCommand((param) =>
            {
                if (string.IsNullOrWhiteSpace(LoginID) || string.IsNullOrWhiteSpace(Password) || string.IsNullOrWhiteSpace(NewPassword))
                {
                    Message = "All fields are mandatory for changing the password.";
                    return;
                }
                //ChangePasswordNow = umgr.ChangePassword(NewPassword, Password, LoginID);
                Message = "";
                if (ChangePasswordNow)
                {
                    Message = "Invalid username or password.";
                }
            });
            ChangePassword = new DelegateCommand((param) =>
            {
                Message = "";
                ChangePasswordNow = true;
            });
            LoginUser = new DelegateCommand((param) =>
            {
                if (string.IsNullOrWhiteSpace(LoginID) || string.IsNullOrWhiteSpace(Password))
                {
                    Message = "UserName and Password both are mandatory for login.";
                    return;
                }
                Models.User user = null;

                string connectionString = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    try
                    {
                        connection.Open();

                        string query = "SELECT LoginId, Password, Role FROM Login WHERE LoginId = @LoginId";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            // Add parameters to prevent SQL injection
                            command.Parameters.AddWithValue("@LoginId", LoginID);
                            command.Parameters.AddWithValue("@Password", Password);

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string dbPassword = reader["Password"].ToString();
                                    string dbRole = reader["Role"].ToString();

                                    if (Password == dbPassword) // Validate password
                                    {
                                        user = new Models.User
                                        {
                                            LoginID = LoginID,
                                            Role = dbRole
                                        };
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Message = $"An error occurred while connecting to the database: {ex.Message}";
                        return;
                    }
                }

                //string Uwd = ConfigurationManager.AppSettings["Uwd"].ToString();
                //string[] users = Uwd.Split('|');
                //foreach (var u in users)
                //{
                //    string[] userDetails = u.Split('$');
                //    if (LoginID.Equals(userDetails[0]) && Password.Equals(userDetails[1]))
                //    {
                //        user = new Models.User();
                //        user.LoginID = userDetails[0];
                //        user.Role = userDetails[2];
                //    }
                //}
                Message = "";
                InCorrectUser = false;
                LoginEventArgs args = new LoginEventArgs();
                if (user != null)
                {
                    args.UserName = LoginID;
                    args.IsValid = true;
                    args.Role = user.Role;
                    args.LoginTime = DateTime.Now;
                    InCorrectUser = false;
                }
                else
                {
                    Message = "Invalid username or password.";
                    InCorrectUser = true;
                    args.IsValid = false;
                }
                LoginID = Password = "";
                OnLogin(args);
            });
        }
    }
}
