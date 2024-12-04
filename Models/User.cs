using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfHelpers;

namespace BaseApp.Models
{
    public class User : ViewModelBase
    {
        private bool _Active;
        private string _Name, _LoginID, _Password, _RoleID, _Role, _FirstLogin;

        public int ID { get; set; }
        public string LastLoginDate { get; set; }
        public string FirstLoginDate { get; set; }

        public string Name
        {
            get => _Name;
            set
            {
                _Name = value;
                OnPropertyChanged("Name");
            }
        }

        public string LoginID
        {
            get => _LoginID;
            set
            {
                _LoginID = value;
                OnPropertyChanged("LoginID");
            }
        }

        public string Password
        {
            get => _Password;
            set
            {
                _Password = value;
                OnPropertyChanged("Password");
            }
        }

        public string RoleID
        {
            get => _RoleID;
            set
            {
                _RoleID = value;
                OnPropertyChanged("RoleID");
            }
        }

        public string Role
        {
            get => _Role;
            set
            {
                _Role = value;
                OnPropertyChanged("Role");
            }
        }

        public string FirstLogin
        {
            get => _FirstLogin;
            set
            {
                _FirstLogin = value;
                OnPropertyChanged("FirstLogin");
            }
        }

        public bool Active
        {
            get => _Active;
            set
            {
                _Active = value;
                OnPropertyChanged("Active");
            }
        }
    }
}
