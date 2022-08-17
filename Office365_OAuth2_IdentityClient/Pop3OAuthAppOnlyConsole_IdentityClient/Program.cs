using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Rebex.Samples;
using Rebex.Net;

namespace Pop3OAuthAppOnlyConsole
{
    /// <summary>
    /// Shows how to authenticate to a mailbox at Office365 (Exchange Online) with OAuth 2.0 (using Microsoft.Identity.Client)
    /// using app-only authentication and retrieve a list of recent mail messages using Rebex Secure Mail (with POP3 protocol).
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
                const int MessageMaxCount = 20;

                // connect using Rebex POP3 and retrieve list of messages
                using (var client = new Pop3())
                {
                    // communication logging (enable if needed)
                    //client.LogWriter = new FileLogWriter("pop3-oauth.log", LogLevel.Debug);

                    // connect to the server
                    Console.WriteLine("Connecting to POP3...");
                    await client.ConnectAsync("outlook.office365.com", SslMode.Implicit);

                    // authenticate using the OAuth 2.0 access token
                    Console.WriteLine("Authenticating to POP3...");
                    await client.LoginAsync(SmtpAddress, accessToken, Pop3Authentication.OAuth20);

                    // list recent messages
                    Console.WriteLine("Listing messages...");
                    int count = client.GetMessageCount();
                    if (count > 0)
                    {
                        for (int num = count; num >= Math.Max(1, count - MessageMaxCount); num--)
                        {
                            Pop3MessageInfo item = await client.GetMessageInfoAsync(num, Pop3ListFields.FullHeaders);
                            Console.WriteLine($"{item.ReceivedDate:yyyy-MM-dd} {item.From}: {item.Subject}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Mailbox is empty.");
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
