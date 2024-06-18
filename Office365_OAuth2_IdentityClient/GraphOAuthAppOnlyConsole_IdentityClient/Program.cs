using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Rebex.Samples;
using Rebex.Net;

namespace GraphOAuthAppOnlyConsole
{
    /// <summary>
    /// Shows how to authenticate to a mailbox with OAuth 2.0 (using Microsoft.Identity.Client)
    /// using app-only authentication and retrieve a list of recent mail messages using Rebex Graph.
    /// See the blog post at https://blog.rebex.net/office365-graph-oauth-unattended for more information.
    /// </summary>
    public static class Program
    {
        //TODO: change the application's client ID, specify client secret value and tenant

        // application's client ID obtained from Azure
        private const string ClientId = "00000000-0000-0000-0000-000000000000"; // configure this

        // application's 'client secret' (can also be referred to as application password
        private const string ClientSecretValue = "ThisIsSomeVerySecretValue"; // configure this

        // your organization's tenant ID
        // (see https://docs.microsoft.com/en-us/azure/active-directory/fundamentals/active-directory-how-to-find-tenant for details)
        private const string TenantId = "00000000-0000-0000-0000-000000000000"; // configure this

        // mailbox to access
        private const string MailAddress = "someone@example.org"; // configure this

        // default scope of permissions to request
        private static readonly string[] Scopes = new[] {
            "https://graph.microsoft.com/.default", // for accessing Exchange Online with app-only auth
        };

        public static async Task Main()
        {
            // get your 30-day trial key at https://www.rebex.net/support/trial/
            Rebex.Licensing.Key = LicenseKey.Value;

            try
            {
                // make sure we have an Azure application client ID and a Rebex key (feel free to remove these checks once configured)
                if (ClientId.Contains("00000000-")) throw new ApplicationException("Please configure ClientId in MainWindow.xaml.cs file.");
                if (string.IsNullOrWhiteSpace(Rebex.Licensing.Key)) throw new ApplicationException("Please set a license key in LicenseKey.cs file.");

                // get an instance of 'CCA' API
                var cca = ConfidentialClientApplicationBuilder
                    .Create(ClientId)
                    .WithClientSecret(ClientSecretValue)
                    .WithTenantId(TenantId)
                    .Build();

                // authenticate interactively for the scopes we need
                Console.WriteLine("Authenticating via Graph...");
                AuthenticationResult result = await cca.AcquireTokenForClient(Scopes).ExecuteAsync();

                // keep the access token and account info
                string accessToken = result.AccessToken;

                // retrieve the list of recent messages
                const int MessageMaxCount = 20;

                // connect using Rebex Graph and retrieve list of messages
                using (var client = new GraphClient())
                {
                    // communication logging (enable if needed)
                    //client.LogWriter = new Rebex.FileLogWriter("graph-oauth.log", Rebex.LogLevel.Debug);

                    // impersonate for a mailbox
                    client.Settings.Impersonation = new GraphImpersonation(MailAddress);

                    // connect to the server
                    Console.WriteLine("Connecting to Exchange Online via Graph API...");
                    await client.ConnectAsync();

                    // authenticate using the OAuth 2.0 access token
                    Console.WriteLine($"Authenticating to Exchange Online via Graph API ({MailAddress})...");
                    await client.LoginAsync(accessToken);

                    // list recent messages in the 'Inbox' folder
                    Console.WriteLine("Listing folder contents...");
                    GraphMessageCollection folderInfo = await client.GetMessageListAsync(GraphFolderId.Inbox, GraphMessageFields.Default, new GraphPageView(0, MessageMaxCount));
                    foreach (var message in folderInfo)
                    {
                        Console.WriteLine($"{message.ReceivedDate:yyyy-MM-dd} {message.From}: {message.Subject}");

                    }
                }

                Console.WriteLine("Finished successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
                return;
            }
        }
    }
}
