using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
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
        //TODO: change the application (client) ID, specify proper tenant and scopes

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

        // credentials that were used to authorize a user
        private OAuthAzureCredentials _credentials;

        // authentication window
        private OAuthAzureAuthorizationWindow _authenticationWindow;

        public MainWindow()
        {
            InitializeComponent();

            // get your 30-day trial key at https://www.rebex.net/support/trial/
            Rebex.Licensing.Key = LicenseKey.Value;
        }

        private void OutlookSign_Click(object sender, RoutedEventArgs e)
        {
            // make sure we have an Azure application client ID and a Rebex key (feel free to remove these checks once configured)
            if (ClientId.Contains("00000000-"))
            {
                MessageBox.Show(this, "Please configure ClientId in MainWindow.xaml.cs file.", "Error");
                return;
            }
            if (string.IsNullOrEmpty(Rebex.Licensing.Key))
            {
                MessageBox.Show(this, "Please set a license key in LicenseKey.cs file.", "Error");
                return;
            }

            // create OAuthOutlookAuthorizationWindow that handles OAuth2 authorization
            statusLabel.Content = "Authenticating via Microsoft 365...";
            _authenticationWindow = new OAuthAzureAuthorizationWindow();
            _authenticationWindow.Owner = this;
            _authenticationWindow.Finished += OutlookSign_Finished;

            // specify the kind of authorization we need
            // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details)
            _authenticationWindow.ClientId = ClientId; // application (client) ID
            _authenticationWindow.TenantId = TenantId; // specify kinds of users to allow
            _authenticationWindow.PromptType = ""; // use default prompt type (also consider "login", "select_account", "consent", ...)
            _authenticationWindow.Scopes = Scopes; // scope of permissions to request

            // start the authentication
            _authenticationWindow.Authorize();
        }

        private void OutlookSign_Finished(object sender, EventArgs e)
        {
            // report error if it occurred
            if (_authenticationWindow.Error != null)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, _authenticationWindow.Error.ToString(), "Error");
                return;
            }

            try
            {
                // obtain the credentials
                _credentials = _authenticationWindow.Credentials;

                // retrieve the list of recent messages
                GetMessageList();
            }
            catch (Exception ex)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, ex.ToString(), "Error");
            }
        }

        private void OutlookRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_credentials == null)
                {
                    statusLabel.Content = "Nothing to refresh. Sign in first.";
                    return;
                }

                statusLabel.Content = "Refreshing...";

                // renew the access token using the refresh token
                _credentials.RefreshAccessToken();

                // retrieve the list of recent messages
                GetMessageList();
            }
            catch (Exception ex)
            {
                statusLabel.Content = "Failure.";
                MessageBox.Show(this, ex.ToString(), "Error");
            }
        }

        private void GetMessageList()
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
                client.Connect("outlook.office365.com", SslMode.Implicit);

                // authenticate using the OAuth 2.0 access token
                statusLabel.Content = "Authenticating to EWS...";
                client.Login(_credentials.AccessToken, EwsAuthentication.OAuth20);

                // list recent messages in the 'Inbox' folder
                statusLabel.Content = "Listing folder contents...";
                EwsFolderInfo folderInfo = client.GetFolderInfo(EwsFolderId.Inbox);
                for (int i = 0; i < Math.Min(MessageMaxCount, folderInfo.ItemsTotal); i += PageSize)
                {
                    EwsMessageCollection list = client.GetMessageList(EwsFolderId.Inbox, EwsItemFields.Envelope, EwsPageView.CreateIndexed(i, Math.Min(PageSize, folderInfo.ItemsTotal - i)));
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
