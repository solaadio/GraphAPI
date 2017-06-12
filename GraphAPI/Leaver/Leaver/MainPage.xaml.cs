using Microsoft.Graph;
using Microsoft.Identity.Client;
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
        /// <summary>
        /// Objects from Microsoft.Client.Identity and Microsoft.Graph to hold data.
        /// Bool variable to check if user has logged in previously. 
        /// </summary>
        GraphServiceClient Client;
        DirectoryObject Manager;
        DirectoryObject Me;
        bool UserExists;
        public MainPage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Check for previous logins and set the variables accordingly.
        /// </summary>
        protected async override void OnAppearing()
        {
            base.OnAppearing();
            if (App.IdentityClientApp.Users.Count() > 0)
            {
                await GetMyDetailsAsync();
                Authenticate.Text = "Sign Out";
                UserExists = true;
            }
            else
            {
                Authenticate.Text = "Sign In";
                UserExists = false;
            }
        }

        /// <summary>
        /// Use this method to get details about signed-in User and his/her Manager
        /// </summary>
        /// <returns></returns>
        private async Task GetMyDetailsAsync()
        {
            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                     new DelegateAuthenticationProvider(
                         async (requestMessage) =>
                         {
                             // var tokenRequest = await App.IdentityClientApp.AcquireTokenAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                             var tokenRequest = await App.IdentityClientApp.AcquireTokenAsync(App.Scopes, App.UiParent);
                             requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                         }));
            Me = await Client.Me.Request().GetAsync();
           Manager = await Client.Me.Manager.Request().GetAsync();
            Username.Text = $"Welcome {((User)Me).DisplayName}";
        }

        /// <summary>
        /// Creates a GraphServiceClient required to call Graph APIs in future
        /// </summary>
        /// <returns>True if client is created successfully. Else False.</returns>
        private async Task<bool> CreateGraphClientAsync()
        {
            try
            {
                Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                     new DelegateAuthenticationProvider(
                         async (requestMessage) =>
                         {
                             var tokenRequest = await App.IdentityClientApp.AcquireTokenAsync(App.Scopes, App.UiParent).ConfigureAwait(false);
                             requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                         }));
                Me = await Client.Me.Request().GetAsync();
                Username.Text = $"Welcome {((User)Me).DisplayName}";
                return true;
            }
            catch (MsalException ex)
            {
                await DisplayAlert("Error", ex.Message, "OK", "Cancel");
                return false;
            }
            
        }

      
        /// <summary>
        /// Authenticates the User and set appropriate variables.
        /// Deletes user details when user is signed-out.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Authenticate_Clicked(object sender, EventArgs e)
        {
            if (!UserExists)
            {
                var status = await CreateGraphClientAsync();
                if (status)
                {
                    await GetMyDetailsAsync();
                    Authenticate.Text = "Signed In";
                    UserExists = true;
                }
            }
            else
            {
                Authenticate.Text = "Signed Out";
                App.IdentityClientApp.Remove(App.IdentityClientApp.Users.FirstOrDefault());
                UserExists = false;
            }
        }

        /// <summary>
        /// Prepares email message with details regarding Full Day leave 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void FullDayLeave_Clicked(object sender, EventArgs e)
        {
            if (!UserExists)
                await CreateGraphClientAsync();
            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                      new DelegateAuthenticationProvider(
                          async (requestMessage) =>
                          {
                              var tokenRequest = await App.IdentityClientApp.AcquireTokenSilentAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                              requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                          }));


            GetMyDetailsAsync();

            var email = new Message
            {
                ToRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Manager).Mail } } },
                CcRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Me).Mail } } },
                Subject = "[Leaver] On Full Day Leave",
                Body = new ItemBody
                {
                    Content = "Hello, <br/>" +
                             "Today, I'll be taking full day leave as I'm not feeling well. For urgent matters please call me on my cell. <br/>" +
                             "Thanks. <br/>" +
                              $"{((User)Me).DisplayName}<br/>" + 
                              $"Sent from { Xamarin.Forms.Device.RuntimePlatform }"                             ,
                    ContentType = BodyType.Html
                }
            };

            SendEmail(email);

        }

         /// <summary>
        /// Prepares email message with details regarding Half Day leave
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void TimeOffForSometimes_Clicked(object sender, EventArgs e)
        {
            if (!UserExists)
                await CreateGraphClientAsync();
            Client = new GraphServiceClient("https://graph.microsoft.com/v1.0",
                      new DelegateAuthenticationProvider(
                          async (requestMessage) =>
                          {
                              var tokenRequest = await App.IdentityClientApp.AcquireTokenSilentAsync(App.Scopes, App.IdentityClientApp.Users.FirstOrDefault());
                              requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenRequest.AccessToken);
                          }));

            GetMyDetailsAsync();

            var email = new Message
            {
                ToRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Manager).Mail } } },
                CcRecipients = new List<Recipient>() { new Recipient() { EmailAddress = new EmailAddress() { Address = ((User)Me).Mail } } },
                Subject = "[Leaver] On Half Day Leave",
                Body = new ItemBody
                {
                    Content = "Hello, <br/>" +
                             "Today, I'll be taking half day leave. For urgent matters please call me on my cell. <br/>" +
                             "Thanks. <br/>" +
                             $"{((User)Me).DisplayName}<br/>" +
                              $"Sent from { Xamarin.Forms.Device.RuntimePlatform }",
                    ContentType = BodyType.Html
                }
            };

            SendEmail(email);
        }

        /// <summary>
        /// Sends an email using Graph APIs
        /// </summary>
        /// <param name="message">Email message to be sent</param>
        private async void SendEmail(Message message)
        {
            if (!UserExists)
                await CreateGraphClientAsync();
            var req = Client.Me.SendMail(message);
            await req.Request().PostAsync();
            Status.Text = $"Email sent to your manager { ((User)Manager).DisplayName }, CC: you";
        }

    }
}
