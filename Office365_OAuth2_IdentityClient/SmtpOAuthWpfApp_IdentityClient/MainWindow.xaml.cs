using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Identity.Client;
#if NET6_0_OR_GREATER
using Microsoft.Identity.Client.Desktop;
#endif
using Rebex.Samples;
using Rebex.Net;

namespace SmtpOAuthWpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// Shows how to authenticate to a mailbox at Microsoft 365 (Office 365, Exchange Online) with OAuth 2.0 (using Microsoft.Identity.Client)
    /// and send an mail messages using Rebex Secure Mail (with SMTP protocol).
    /// See the blog post at https://blog.rebex.net/oauth2-office365-rebex-mail for more information.
    /// </summary>
    public partial class MainWindow : Window
    {
        //TODO: change the application (client) ID, specify proper tenant, scopes and prompt type

        // application (client) ID obtained from Azure
        private const string ClientId = "00000000-0000-0000-0000-000000000000";

        // specifies which users to allow (also consider "common", "consumers", domain name or a GUID identifier)
        // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols#endpoints for details)
        private const string TenantId = "organizations";
                
        // scope of permissions to request from the user
        // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent for details)
        private static readonly string[] Scopes = new[] {
            "profile", // needed to retrieve the user name, which is required for Office 365's SMTP authentication
            "email", // not required, but may be useful
            "openid", // required by the 'profile' and 'email' scopes
            "offline_access", // specify this scope to make it possible to refresh the access token when it expires (after one hour)
            "https://outlook.office365.com/SMTP.Send", // scope for sending Microsoft 365 Exchange Online mail via SMTP
        };

        // specifies the kind of login/consent dialog
        private static readonly Prompt PromptType = Prompt.SelectAccount;


        // 'PCA' API
        private IPublicClientApplication _publicClientApplication;

        // OAuth 2.0 access token
        private string _accessToken;

        // user account
        private IAccount _account;

        public MainWindow()
        {
            InitializeComponent();

            // get your 30-day trial key at https://www.rebex.net/support/trial/
            Rebex.Licensing.Key = LicenseKey.Value;
        }

        private async void OutlookSign_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // make sure we have an Azure application client ID and a Rebex key (feel free to remove these checks once configured)
                if (ClientId.Contains("00000000-")) throw new ApplicationException("Please configure ClientId in MainWindow.xaml.cs file.");
                if (string.IsNullOrWhiteSpace(Rebex.Licensing.Key)) throw new ApplicationException("Please set a license key in LicenseKey.cs file.");

                // specify options
                var options = new PublicClientApplicationOptions()
                {
                    ClientId = ClientId,
                    TenantId = TenantId,
                    // redirect URI (this one is suitable for desktop and mobile apps with embedded browser)
                    RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",
                };

                // get an instance of 'PCA' API
                _publicClientApplication = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(options)
                    .WithParentActivityOrWindow(() => new System.Windows.Interop.WindowInteropHelper(this).Handle)
#if NET6_0_OR_GREATER
                    .WithWindowsEmbeddedBrowserSupport()
#endif
                    .Build();

                // authenticate interactively for the scopes we need
                statusLabel.Content = "Authenticating via Microsoft 365...";
                AuthenticationResult result = await _publicClientApplication
                    .AcquireTokenInteractive(Scopes)
                    .WithPrompt(PromptType)
#if NET6_0_OR_GREATER
                    .WithUseEmbeddedWebView(true)
#endif
                    .ExecuteAsync();

                // keep the access token and account info
                _accessToken = result.AccessToken;
                _account = result.Account;
                // Note: In real applications, you would most likely want to keep result.ExpiresOn as well

                // make sure we obtained the user name
                string userName = _account.Username;
                if (userName == null)
                {
                    throw new InvalidOperationException("User name not retrieved. Make sure you specified 'openid' and 'profile' scopes.");
                }

                // connect and authenticate to the server to ensure it works
                using (var client = await ConnectAndAuthenticateAsync())
                {
                }

                statusLabel.Content = "Successfully authenticated.";
            }
            catch (Exception ex)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, ex.ToString(), "Error");
            }
        }

        private async void OutlookRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_accessToken == null || _account == null)
                {
                    statusLabel.Content = "Nothing to refresh. Please sign in first.";
                    return;
                }

                statusLabel.Content = "Refreshing...";

                // authenticate silently for the scopes we need (refreshing the token if needed)
                AuthenticationResult result = await _publicClientApplication.AcquireTokenSilent(Scopes, _account).ExecuteAsync();

                // update the access token
                _accessToken = result.AccessToken;
                // Note: In real applications, you would most likely want to keep result.ExpiresOn as well

                // connect and authenticate to the server to ensure it works
                using (var client = await ConnectAndAuthenticateAsync())
                {
                }

                statusLabel.Content = "Token refreshed successfully.";
            }
            catch (Exception ex)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, ex.ToString(), "Error");
            }
        }

        private async Task<Smtp> ConnectAndAuthenticateAsync()
        {
            // connect and authenticate to the server and returns an instance of Smtp class
            var client = new Smtp();
            try
            {
                // communication logging (enable if needed)
                //client.LogWriter = new Rebex.FileLogWriter("smtp-oauth.log", Rebex.LogLevel.Debug);

                await client.ConnectAsync("smtp.office365.com", SslMode.Explicit);
                await client.LoginAsync(_account.Username, _accessToken, SmtpAuthentication.OAuth20);
                return client;
            }
            catch
            {
                client.Dispose();
                throw;
            }
        }

        private async void sendButton_Click(object sender, RoutedEventArgs e)
        {
            if (_accessToken == null)
            {
                MessageBox.Show(this, "Please sign-in first to retrieve an access token.", "Error");
                return;
            }

            try
            {
                // establish authenticated SMTP connection
                using (Smtp smtp = await ConnectAndAuthenticateAsync())
                {
                    // send the message
                    smtp.Send(fromTextBox.Text, toTextBox.Text, subjectTextBox.Text, bodyTextBox.Text);

                    statusLabel.Content = "Message sent!";
                }
            }
            catch (Exception ex)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, ex.ToString(), "Error");
            }
        }
    }
}
