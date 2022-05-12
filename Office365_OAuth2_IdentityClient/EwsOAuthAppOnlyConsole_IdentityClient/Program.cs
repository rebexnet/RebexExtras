using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Rebex.Samples;
using Rebex.Net;

namespace EwsOAuthAppOnlyConsole
{
    /// <summary>
    /// Shows how to authenticate to a mailbox at Office365 (Exchange Online) with OAuth 2.0 (using Microsoft.Identity.Client)
    /// using app-only authentication and retrieve a list of recent mail messages using Rebex Secure Mail (with EWS protocol).
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
        private const string SmtpAddress = "someone@example.org"; // configure this

        // default scope of permissions to request
        private static readonly string[] Scopes = new[] {
            "https://outlook.office365.com/.default", // scope for accessing Office365 via Exhange Web Services with app-only auth
        };

        public static async Task Main()
        {
            // get your 30-day trial key at https://www.rebex.net/support/trial/
            Rebex.Licensing.Key = LicenseKey.Value;

            try
            {
                // make sure we have an Azure application client ID and a Rebex key (feel free to remove these checks once configured)
                if (ClientId.Contains("00000000-")) throw new ApplicationException("Please configure ClientId in MainWindow.xaml.cs file.");
                if (Rebex.Licensing.Key.Contains("_TRIAL_KEY_")) throw new ApplicationException("Please set a license key in LicenseKey.cs file.");

                // get an instanca of 'CCA' API
                var cca = ConfidentialClientApplicationBuilder
                    .Create(ClientId)
                    .WithClientSecret(ClientSecretValue)
                    .WithTenantId(TenantId)
                    .Build();

                // authenticate interactively for the scopes we need
                Console.WriteLine("Authenticating via Office365...");
                AuthenticationResult result = await cca.AcquireTokenForClient(Scopes).ExecuteAsync();

                // keep the access token and account info
                string accessToken = result.AccessToken;

                // retrieve the list of recent messages
                const int PageSize = 5;
                const int MessageMaxCount = 20;

                // connect using Rebex EWS and retrieve list of messages
                using (var client = new Ews())
                {
                    // communication logging (enable if needed)
                    //client.LogWriter = new FileLogWriter("ews-oauth.log", LogLevel.Debug);

                    /// impersonate for a mailbox
                    client.Settings.Impersonation = new EwsImpersonation() { SmtpAddress = SmtpAddress };

                    // connect to the server
                    Console.WriteLine("Connecting to EWS...");
                    await client.ConnectAsync("outlook.office365.com", SslMode.Implicit);

                    // authenticate using the OAuth 2.0 access token
                    Console.WriteLine("Authenticating to EWS...");
                    await client.LoginAsync(accessToken, EwsAuthentication.OAuth20);

                    // list recent messages in the 'Inbox' folder
                    Console.WriteLine("Listing folder contents...");
                    EwsFolderInfo folderInfo = await client.GetFolderInfoAsync(EwsFolderId.Inbox);
                    for (int i = 0; i < Math.Min(MessageMaxCount, folderInfo.ItemsTotal); i += PageSize)
                    {
                        EwsMessageCollection list = await client.GetMessageListAsync(EwsFolderId.Inbox, EwsItemFields.Envelope, EwsPageView.CreateIndexed(i, Math.Min(PageSize, folderInfo.ItemsTotal - i)));
                        foreach (EwsMessageInfo item in list)
                        {
                            Console.WriteLine($"{item.ReceivedDate:yyyy-MM-dd} {item.From}: {item.Subject}");
                        }
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
