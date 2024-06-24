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
using Rebex.Samples;
using Rebex.Net;

namespace EwsOAuthWpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// Shows how to authenticate to a mailbox at Microsoft 365 (Office 365, Exchange Online) with OAuth 2.0
    /// and retrieve a list of recent mail messages using Rebex Secure Mail (with EWS protocol).
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
            //"profile", // needed to retrieve the user name (not required by EWS, but may be useful)
            //"email", // not required, but may be useful
            //"openid", // required by the 'profile' and 'email' scopes
            "offline_access", // specify this scope to make it possible to refresh the access token when it expires (after one hour)
            "https://outlook.office365.com/EWS.AccessAsUser.All", // scope for accessing Microsoft 365 Exchange Online via EWS
        };

        // specifies appearance of the login/consent dialog
        private const string PromptType = OAuthPromptType.SelectAccount;


        // credentials that were used to authorize a user
        private OAuthAzureCredentials _credentials;

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

                // create OAuthOutlookAuthorizationWindow that handles OAuth2 authorization
                statusLabel.Content = "Authenticating via Microsoft 365...";
                var authenticationWindow = new OAuthAzureAuthorizationWindow();
                authenticationWindow.Owner = this;

                // specify the kind of authorization we need
                // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details)
                authenticationWindow.ClientId = ClientId; // application (client) ID
                authenticationWindow.TenantId = TenantId; // specify kinds of users to allow
                authenticationWindow.PromptType = PromptType; // appearance of the login/consent dialog
                authenticationWindow.Scopes = Scopes; // scope of permissions to request

                // perform the authorization and obtain the credentials
                _credentials = await authenticationWindow.AuthorizeAsync();

                // retrieve the list of recent messages
                await GetMessageListAsync();
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
                if (_credentials == null)
                {
                    statusLabel.Content = "Nothing to refresh. Please sign in first.";
                    return;
                }

                statusLabel.Content = "Refreshing...";

                // renew the access token using the refresh token
                await _credentials.RefreshAccessTokenAsync();

                // retrieve the list of recent messages
                await GetMessageListAsync();
            }
            catch (Exception ex)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, ex.ToString(), "Error");
            }
        }

        private async Task GetMessageListAsync()
        {
            const int PageSize = 5;
            const int MessageMaxCount = 20;

            lvItems.Items.Clear();

            // connect using Rebex EWS and retrieve list of messages
            using (var client = new Ews())
            {
                // communication logging (enable if needed)
                //client.LogWriter = new Rebex.FileLogWriter("ews-oauth.log", Rebex.LogLevel.Debug);

                // connect to the server
                statusLabel.Content = "Connecting to EWS...";
                await client.ConnectAsync("outlook.office365.com", SslMode.Implicit);

                // authenticate using the OAuth 2.0 access token
                statusLabel.Content = "Authenticating to EWS...";
                await client.LoginAsync(_credentials.AccessToken, EwsAuthentication.OAuth20);

                // list recent messages in the 'Inbox' folder
                statusLabel.Content = "Listing folder contents...";
                EwsFolderInfo folderInfo = await client.GetFolderInfoAsync(EwsFolderId.Inbox);
                for (int i = 0; i < Math.Min(MessageMaxCount, folderInfo.ItemsTotal); i += PageSize)
                {
                    EwsMessageCollection list = await client.GetMessageListAsync(EwsFolderId.Inbox, EwsItemFields.Envelope, EwsPageView.CreateIndexed(i, Math.Min(PageSize, folderInfo.ItemsTotal - i)));
                    foreach (EwsMessageInfo item in list)
                    {
                        lvItems.Items.Add($"{item.ReceivedDate:yyyy-MM-dd} {item.From}: {item.Subject}");
                    }
                }
            }

            statusLabel.Content = "Finished successfully!";
        }
    }
}
