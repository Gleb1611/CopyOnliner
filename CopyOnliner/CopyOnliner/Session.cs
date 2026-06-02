using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CopyOnliner
{
    public static class Session
    {
        public static User CurrentUser { get; set; }

        public static bool IsLoggedIn
        {
            get { return CurrentUser != null && CurrentUser.IsLoggedIn; }
        }

        public static void Login(User user)
        {
            CurrentUser = user;
        }

        public static void Logout()
        {
            CurrentUser = null;
        }
    }
}
