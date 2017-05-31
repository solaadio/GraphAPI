using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Leaver
{

    public partial class MainPage : ContentPage
    {
        GraphServiceClient client;
        DirectoryObject Manager;
        DirectoryObject Me;
        bool UserExists;
        public MainPage()
        {
            InitializeComponent();
        }

        protected override void OnAppearing()
        {
            base.OnAppearing();
            if (App.IdentityClientApp.Users.Count() > 0)
            {
                Authenticate.Text = "Sign Out";
                UserExists = true;
            }
            else
            {
                Authenticate.Text = "Sign In";
                UserExists = false;
            }
        }
        private async Task CreateGraphClient()
        {
            client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                     new DelegateAuthenticationProvider(
                         async (requestMessage) =>
                         {
                             var tokenRequest = await App.IdentityClientApp.AcquireTokenAsync(App.Scopes, App.UiParent).ConfigureAwait(false);
                             requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                         }));
            Me = await client.Me.Request().GetAsync();
            Username.Text = $"Welcome {((User)Me).DisplayName}";
        }

        private async void Authenticate_Clicked(object sender, EventArgs e)
        {
            if (!UserExists)
            {
                await CreateGraphClient();
                Authenticate.Text = "Signed In";
                UserExists = true;
            }
            else
            {
                Authenticate.Text = "Sign Out";
                App.IdentityClientApp.Remove(App.IdentityClientApp.Users.FirstOrDefault());
                UserExists = false;
            }
        }

        private async void FullDayLeave_Clicked(object sender, EventArgs e)
        {
            if (!UserExists)
                await CreateGraphClient();
            client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                      new DelegateAuthenticationProvider(
                          async (requestMessage) =>
                          {
                              var tokenRequest = await App.IdentityClientApp.AcquireTokenSilentAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                              requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                          }));

            Manager = await client.Me.Manager.Request().GetAsync();
            Me = await client.Me.Request().GetAsync();

            var email = new Message
            {
                ToRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Me).Mail } } },
                CcRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Me).Mail } } },
                Subject = "[Leaver] On Half Day Leave",
                Body = new ItemBody
                {
                    Content = "Hello, <br/>" +
                             "Today, I'll be working half day today. For urgent matters please call me on my cell. <br/>" +
                             "Thanks. <br/>" +
                             ((User)Me).DisplayName,
                    ContentType = BodyType.Html
                }
            };

            SendEmail(email);

        }

     
        private async void TimeOffForSometimes_Clicked(object sender, EventArgs e)
        {
            if (!UserExists)
                CreateGraphClient();
            client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                      new DelegateAuthenticationProvider(
                          async (requestMessage) =>
                          {
                              var tokenRequest = await App.IdentityClientApp.AcquireTokenSilentAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                              requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                          }));

            Manager = await client.Me.Manager.Request().GetAsync();
            Me = await client.Me.Request().GetAsync();

            var email = new Message
            {
                ToRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Me).Mail } } },
                CcRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Me).Mail } } },
                Subject = "[Leaver] On Half Day Leave",
                Body = new ItemBody
                {
                    Content = "Hello, <br/>" +
                             "Today, I'll be taking full day leave as I'm not feeling well. For urgent matters please call me on my cell. <br/>" +
                             "Thanks. <br/>" +
                             ((User)Me).DisplayName,
                    ContentType = BodyType.Html
                }
            };

            SendEmail(email);
        }

        private async void SendEmail(Message message)
        {
            if (!UserExists)
                CreateGraphClient();
            var req = client.Me.SendMail(message);
            await req.Request().PostAsync();
            Status.Text = "Email sent to your manager, CC: you";
        }

    }
}
