using System;

namespace CopyOnliner
{
    public class User
    {
        public int UserID { get; set; }
        public string Username { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public bool IsLoggedIn { get; set; }
    }
}