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

namespace ImapOAuthWpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// Shows how to authenticate to a mailbox at Office365 (Exchange Online) with OAuth 2.0
    /// and retrieve a list of recent mail messages using Rebex Secure Mail (with IMAP protocol).
    /// See the blog post at https://blog.rebex.net/oauth2-office365-rebex-mail for more information.
    /// </summary>
    public partial class MainWindow : Window
    {
        //TODO: change the application's client ID, specify proper tenant and scopes

        // application's client ID obtained from Azure
        private const string ClientId = "00000000-0000-0000-0000-000000000000";

        // specifies which users to allow (also consider "common", "consumers", domain name or a GUID identifier)
        // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols#endpoints for details)
        private const string TenantId = "organizations";

        // scope of permissions to request from the user
        // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent for details)
        private static readonly string[] Scopes = new[] {
            "profile", // needed to retrieve the user name, which is required for Office365's IMAP authentication
            //"email", // not required, but may be useful
            "openid", // required by the 'profile' and 'email' scopes
            "offline_access", // specify this scope to make it possible to refresh the access token when it expires (after one hour)
            "https://outlook.office365.com/IMAP.AccessAsUser.All", // scope for accessing Office365 via IMAP
        };

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
                if (Rebex.Licensing.Key.Contains("_TRIAL_KEY_")) throw new ApplicationException("Please set a license key in LicenseKey.cs file.");

                // create OAuthOutlookAuthorizationWindow that handles OAuth2 authorization
                statusLabel.Content = "Authenticating via Office365...";
                var authenticationWindow = new OAuthAzureAuthorizationWindow();
                authenticationWindow.Owner = this;

                // specify the kind of authorization we need
                // (see https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details)
                authenticationWindow.ClientId = ClientId; // application's client ID
                authenticationWindow.TenantId = TenantId; // specify kinds of users to allow
                authenticationWindow.PromptType = ""; // use default prompt type (also consider "login", "select_account", "consent", ...)
                authenticationWindow.Scopes = Scopes; // scope of permissions to request

                // perform the authorization and obtain the credentials
                var credentials = await authenticationWindow.AuthorizeAsync();

                // make sure we obtained the user name
                string userName = credentials.UserName;
                if (userName == null)
                {
                    throw new InvalidOperationException("User name not retrieved. Make sure you specified 'openid' and 'profile' scopes.");
                }

                // keep the credentials
                _credentials = credentials;

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
                    statusLabel.Content = "Nothing to refresh. Sign in first.";
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
            lvItems.Items.Clear();

            // connect using Rebex IMAP and retrieve list of messages
            using (var client = new Imap())
            {
                // communication logging (enable if needed)
                // client.LogWriter = new FileLogWriter("imap-oauth.log", LogLevel.Debug);

                // connect to the server
                statusLabel.Content = "Connecting to IMAP...";
                await client.ConnectAsync("outlook.office365.com", Imap.DefaultImplicitSslPort, SslMode.Implicit);

                // prepare authentication token suitable for IMAP, POP3, or SMTP
                string userName = _credentials.UserName;
                string accessToken = _credentials.AccessToken;
                string pattern = string.Format("user={0}{1}auth=Bearer {2}{1}{1}", userName, '\x1', accessToken);
                string token = Convert.ToBase64String(Encoding.ASCII.GetBytes(pattern));

                // authenticate using the OAuth 2.0 access token
                statusLabel.Content = "Authenticating to IMAP...";
                await client.LoginAsync(token, ImapAuthentication.OAuth20);

                // list recent messages in the 'Inbox' folder

                statusLabel.Content = "Listing folder contents...";
                await client.SelectFolderAsync("Inbox", readOnly: true);

                int messageCount = client.CurrentFolder.TotalMessageCount;
                var messageSet = new ImapMessageSet();
                messageSet.AddRange(Math.Max(1, messageCount - 50), messageCount);

                var list = await client.GetMessageListAsync(messageSet, ImapListFields.Envelope);
                list.Sort(new ImapMessageInfoComparer(ImapMessageInfoComparerType.SequenceNumber, Rebex.SortingOrder.Descending));
                foreach (ImapMessageInfo item in list)
                {
                    lvItems.Items.Add($"{item.Date.LocalTime:yyyy-MM-dd} {item.From}: {item.Subject}");
                }
            }

            statusLabel.Content = "Finished successfully!";
        }
    }
}
