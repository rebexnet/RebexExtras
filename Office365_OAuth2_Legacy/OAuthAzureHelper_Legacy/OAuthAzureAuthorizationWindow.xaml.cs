using System;
using System.Windows;

namespace Rebex.Samples
{
    /// <summary>
    /// Implements OAuth2 authentication logic for OAuthOutlookAuthorizationWindow.xaml window.
    /// </summary>
    public partial class OAuthAzureAuthorizationWindow : Window
    {
        /// <summary>
        /// Application (client) ID. Assigned by Azure portal's App registrations.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Controls who can sign into the application. Allowed values include 'common', 'organizations', 'consumers', domain, or a GUID identifier.
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols#endpoints for details.
        /// </summary>
        public string TenantId { get; set; } = "organizations";

        /// <summary>
        /// Type of required user interaction.
        /// "" = default (don't request credentials if already signed-on); 
        /// "login" = always request credentials;
        /// "none" = never request credential;
        /// "consent" = ask the user to grant permissions to the app;
        /// "select_account" = ask the user to select the account to use.
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details.
        /// </summary>
        public string PromptType { get; set; } = string.Empty;

        /// <summary>
        /// The redirect URI of your application. Used to retrieve authentication responses.
        /// Must exactly match one of the redirect_uris registered with the application.
        /// For desktop and mobile applications, https://login.microsoftonline.com/common/oauth2/nativeclient is used
        /// (which is also the default value used by this class when the argument is not specified).
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details.
        /// </summary>
        public string RedirectUri { get; set; }

        /// <summary>
        /// List of scopes we want the user to consent to.
        /// "openid" = needed to retrieve user info (required for "email" and "profile" scopes);
        /// "profile" = retrieve the username and full name (this is a must for IMAP);
        /// "email" = retrieve the user's e-mail address;
        /// "offline_access" = makes it possible for the application to refresh the access token when it expires;
        /// "https://outlook.office365.com/EWS.AccessAsUser.All" = for accessing Microsoft 365 via EWS;
        /// "https://outlook.office365.com/IMAP.AccessAsUser.All" = for accessing Microsoft 365 via IMAP.
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent for details.
        /// </summary>
        public string[] Scopes { get; set; } = new[] { "openid", "profile" };

        /// <summary>
        /// Creates an instance of <see cref="OAuthAzureAuthorizationWindow"/>.
        /// </summary>
        public OAuthAzureAuthorizationWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Event handler that is raised when the authentication has finished.
        /// </summary>
        public event EventHandler Finished;

        /// <summary>
        /// Class for OAuth 2.0 credentials that encapsulates most of the authorization logic as well.
        /// </summary>
        public OAuthAzureCredentials Credentials { get; private set; }

        /// <summary>
        /// Error that occurred during the authentication, or null on success.
        /// </summary>
        public Exception Error { get; private set; }

        /// <summary>
        /// Indicates whether authentication is in progress.
        /// </summary>
        private bool _authenticating;

        /// <summary>
        /// Performs OAuth 2.0 authorization.
        /// Opens a window with browser control to make it possible for users to supply their username and password to Microsoft's authentication servers.
        /// </summary>
        public void Authorize()
        {
            if (ClientId == null)
            {
                throw new InvalidOperationException("ClientId not specified.");
            }

            if (TenantId == null)
            {
                throw new InvalidOperationException("TenantId not specified.");
            }

            if (Scopes == null)
            {
                throw new InvalidOperationException("Scopes not specified.");
            }

            if (Credentials != null)
            {
                throw new InvalidOperationException("Only one authentication request can be performed.");
            }

            // get redirect URI and determine expected authority
            string redirectUri = RedirectUri ?? OAuthAzureCredentials.DefaultRedirectUri;

            // create an instance of OAuthAzureCredentials helper class
            Credentials = new OAuthAzureCredentials(ClientId, TenantId, PromptType, redirectUri, Scopes);

            // set field that indicates that authentication is in progress
            _authenticating = true;

            // Direct the user to authorization endpoint.
            // Once completed, the application will receive an authorization code.

            // show the window and navigate to authentication URI
            Show();
            webBrowser.Navigate(new Uri(Credentials.AuthorizationUri));
        }

        /// <summary>
        /// Makes sure we landed on an expected authentication URI.
        /// </summary>
        /// <param name="uri">URI.</param>
        /// <returns>True if OK.</returns>
        private bool CheckExpectedAuthority(Uri uri)
        {
            // abort authentication if we navigated to an unexpected domain
            // (this usually means that Microsoft's authentication website is not accessible)
            var expectedAuthority = new Uri(Credentials.RedirectUri).GetLeftPart(UriPartial.Authority);
            if (uri.GetLeftPart(UriPartial.Authority) != expectedAuthority)
            {
                Close();
                if (_authenticating)
                {
                    Error = new OAuthAzureException("Unable to open authorization URL.");
                    Finished?.Invoke(this, EventArgs.Empty);
                    _authenticating = false;
                }
                return false;
            }

            return true;
        }

        /// <summary>
        /// Handler of browser control navigated event.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Navigation event args.</param>
        private void WebBrowser_Navigated(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            CheckExpectedAuthority(e.Uri);
        }

        /// <summary>
        /// Handler of the browser control load completed event.
        /// If the URI contains the authorization code, proceed to obtain the user's OAuth2 token.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Navigation event args.</param>
        private void WebBrowser_LoadCompleted(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            if (!_authenticating)
            {
                // no need to do anything here if we did not navigate to the proper URI
                return;
            }

            if (!CheckExpectedAuthority(e.Uri))
            {
                return;
            }

            try
            {
                // get query part of the browser URL
                var query = System.Web.HttpUtility.ParseQueryString(e.Uri.Query);

                // get error message from the supplied URL
                string errorCode = query["error"];
                if (errorCode != null)
                {
                    string description = query["error_description"] ?? string.Concat("Error '", errorCode, "'.");
                    throw new OAuthAzureException(description, errorCode);
                }

                // get authorization code from the supplied URL
                string authorizationCode = query["code"];
                if (authorizationCode == null)
                {
                    // we are not on the final authentication page yet...
                    return;
                }

                // The user has granted us permissions, and we received an authorization code.
                // Next, we will exchange the code for an access token using the '/token' endpoint.

                Credentials.RequestAccessToken(authorizationCode);

                // finish the successful asynchronous authentication request and close the window
                if (_authenticating)
                {
                    Finished?.Invoke(this, EventArgs.Empty);
                    _authenticating = false;
                }
                Close();
            }
            catch (Exception error)
            {
                if (error is OAuthAzureException)
                {
                    Error = error;
                }
                else
                {
                    Error = new OAuthAzureException("Error during OAuth2 authentication. " + error.Message, error);
                }
                Close();
                if (_authenticating)
                {
                    Error = new OAuthAzureException("Unable to open authorization URL.");
                    Finished?.Invoke(this, EventArgs.Empty);
                    _authenticating = false;
                }
            }
        }

        /// <summary>
        /// Handles premature windows closure.
        /// </summary>
        /// <param name="sender">Sender.</param>
        /// <param name="e">Event args.</param>
        private void Window_Closed(object sender, EventArgs e)
        {
            if (_authenticating)
            {
                Error = Error ?? new OAuthAzureException("Authentication window has been closed unexpectedly.");
                Finished?.Invoke(this, EventArgs.Empty);
                _authenticating = false;
            }
        }
    }
}
