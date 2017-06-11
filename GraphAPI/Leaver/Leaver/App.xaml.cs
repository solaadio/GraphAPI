using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Xamarin.Forms;

namespace Leaver
{
    public partial class App : Application
    {
        public static PublicClientApplication IdentityClientApp = null;
        public static string ClientID = "9059082d-b0f4-4cdd-aab3-079fa3a906b8";
        public static string[] Scopes = { "User.Read", "User.ReadBasic.All ", "Mail.Send" };
        public static UIParent UiParent = null;
        public App()
        {
            InitializeComponent();
            IdentityClientApp = new PublicClientApplication(ClientID);
            MainPage = new Leaver.MainPage();
        }

        protected override void OnStart()
        {
            // Handle when your app starts
        }

        protected override void OnSleep()
        {
            // Handle when your app sleeps
        }

        protected override void OnResume()
        {
            // Handle when your app resumes
        }
    }
}
