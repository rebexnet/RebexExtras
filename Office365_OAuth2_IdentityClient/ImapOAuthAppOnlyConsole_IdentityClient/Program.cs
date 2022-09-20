using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Rebex.Samples;
using Rebex.Net;

namespace ImapOAuthAppOnlyConsole
{
    /// <summary>
    /// Shows how to authenticate to a mailbox at Microsoft 365 (Office 365, Exchange Online) with OAuth 2.0 (using Microsoft.Identity.Client)
    /// using app-only authentication and retrieve a list of recent mail messages using Rebex Secure Mail (with IMAP protocol).
    /// </summary>
    public static class Program
    {
        //TODO: change the application (client) ID, specify client secret value and tenant

        // application (client) ID obtained from Azure
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
            "https://outlook.office365.com/.default", // scope for accessing Microsoft 365 Exchange Online with app-only auth
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

                // get an instance of 'CCA' API
                var cca = ConfidentialClientApplicationBuilder
                    .Create(ClientId)
                    .WithClientSecret(ClientSecretValue)
                    .WithTenantId(TenantId)
                    .Build();

                // authenticate interactively for the scopes we need
                Console.WriteLine("Authenticating via Microsoft 365...");
                AuthenticationResult result = await cca.AcquireTokenForClient(Scopes).ExecuteAsync();

                // keep the access token and account info
                string accessToken = result.AccessToken;

                // retrieve the list of recent messages
                const int MessageMaxCount = 20;

                // connect using Rebex IMAP and retrieve list of messages
                using (var client = new Imap())
                {
                    // communication logging (enable if needed)
                    //client.LogWriter = new Rebex.FileLogWriter("imap-oauth.log", Rebex.LogLevel.Debug);

                    // connect to the server
                    Console.WriteLine("Connecting to IMAP...");
                    await client.ConnectAsync("outlook.office365.com", SslMode.Implicit);

                    // authenticate using the OAuth 2.0 access token
                    Console.WriteLine($"Authenticating to IMAP ({SmtpAddress})...");
                    await client.LoginAsync(SmtpAddress, accessToken, ImapAuthentication.OAuth20);

                    // list recent messages in the 'Inbox' folder
                    Console.WriteLine("Listing folder contents...");
                    await client.SelectFolderAsync("Inbox", readOnly: true);
                    int count = client.CurrentFolder.TotalMessageCount;
                    if (count > 0)
                    {
                        var range = new ImapMessageSet();
                        range.AddRange(Math.Max(1, count - MessageMaxCount), count);
                        var list = await client.GetMessageListAsync(range, ImapListFields.Envelope);
                        foreach (ImapMessageInfo item in list)
                        {
                            Console.WriteLine($"{item.ReceivedDate:yyyy-MM-dd} {item.From}: {item.Subject}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Folder '{0}' is empty.", client.CurrentFolder.Name);
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
