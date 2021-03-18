using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Rebex.Samples
{
    /// <summary>
    /// Helper class for retrieval of OAuth2 credentials suitable for Office365 from Microsoft via Azure.
    /// Implements the authorization code flow described at
    /// https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code
    /// using .NET API and System.Text.Json (or Newtonsoft.Json).
    /// </summary>
    /// <remarks>
    /// See the blog post at https://blog.rebex.net/oauth2-office365-rebex-mail for more information.
    /// </remarks>
    public class OAuthAzureCredentials
    {
        /// <summary>
        /// Access token. Expires in an hour.
        /// If 'offline_access' scope was specified, <see cref="RefreshAccessTokenAsync"/> method can be used to obtain a new one.
        /// </summary>
        public string AccessToken { get; private set; }

        /// <summary>
        /// User name of the authenticated user.
        /// Only available if 'profile' and 'openid' scopes were specified.
        /// </summary>
        public string UserName { get; private set; }

        /// <summary>
        /// Full Name of the authenticated user.
        /// Only available if 'profile' and 'openid' scopes were specified.
        /// </summary>
        public string FullName { get; private set; }

        /// <summary>
        /// E-mail of the authenticated user.
        /// Only available if 'email' and 'openid' scopes were specified.
        /// </summary>
        public string Email { get; private set; }

        /// <summary>
        /// Refresh token. Used to request the access token when it expires.
        /// This is used used by <see cref="RefreshAccessTokenAsync"/> method to obtain a new <see cref="AccessToken"/> when the old one expires.
        /// </summary>
        public string RefreshToken { get; private set; }

        /// <summary>
        /// Authorization URI. This is an URL at Microsoft's servers where the user gets directed to for authorization.
        /// </summary>
        public string AuthorizationUri { get; private set; }

        /// <summary>
        /// Redirection URI. This is where the user's browser will be redirected once authentication is over.
        /// </summary>
        public string RedirectUri { get; private set; }

        /// <summary>
        /// Redirect URI for use with desktop and mobile applications with embedded browsers.
        /// </summary>
        public const string DefaultRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient";

        /// <summary>
        /// Creates an instance of <see cref="OAuthAzureCredentials"/>.
        /// </summary>
        /// <param name="clientId">
        /// Application (client) ID. Assigned by Azure portal's App registrations.
        /// </param>
        /// <param name="tenantId">
        /// Controls who can sign into the application. Allowed valules include 'common', 'organizations', 'consumers', domain, or a GUID identifier.
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols#endpoints for details.
        /// </param>
        /// <param name="promptType">
        /// Type of required user interaction.
        /// "" = default (don't request credentials if already signed-on); 
        /// "login" = always request credentials;
        /// "none" = never request credential;
        /// "consent" = ask the user to grant permissions to the app;
        /// "select_account" = ask the user to select the account to use.
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details.
        /// </param>
        /// <param name="redirectUri">
        /// The redirect URI of your application. Used to retrieve authentication responses.
        /// Must exactly match one of the redirect_uris registered with the application.
        /// For desktop and mobile applications, https://login.microsoftonline.com/common/oauth2/nativeclient is used
        /// (which is also the default value used by this class when the argument is not specified).
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details.
        /// </param>
        /// <param name="scopes">
        /// List of scopes we want the user to consent to.
        /// "openid" = needed to retrieve user info (required for "email" and "profile" scopes);
        /// "profile" = retrieve the username and full name (this is a must for IMAP);
        /// "email" = retrieve the user's e-mail address;
        /// "offline_access" = makes it possible for the application to refresh the access token when it expires;
        /// "https://outlook.office365.com/EWS.AccessAsUser.All" = for accessing Office365 via EWS;
        /// "https://outlook.office365.com/IMAP.AccessAsUser.All" = for accessing Office365 via IMAP.
        /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent for details.
        /// </param>
        public OAuthAzureCredentials(string clientId, string tenantId, string promptType, string redirectUri, params string[] scopes)
        {
            if (clientId == null)
            {
                throw new ArgumentNullException("clientId");
            }

            if (tenantId == null)
            {
                throw new ArgumentNullException("tenantId");
            }

            if (scopes == null)
            {
                throw new ArgumentNullException("scopes");
            }

            _clientId = clientId;
            _scopes = Uri.EscapeDataString(string.Join(" ", scopes));

            _tokenEndPoint = string.Format(CultureInfo.InvariantCulture, AzureEndPoint, tenantId, "token");
            _authorizeEndPoint = string.Format(CultureInfo.InvariantCulture, AzureEndPoint, tenantId, "authorize");

            promptType = promptType ?? string.Empty; // use default prompt type if not specified
            RedirectUri = redirectUri ?? DefaultRedirectUri; // use default redirect URI for desktop and mobile application if not specified

            AuthorizationUri = string.Format(
                CultureInfo.InvariantCulture,
                "{0}?client_id={1}&response_type=code&response_mode=query&redirect_uri={2}&scope={3}&prompt={4}",
                _authorizeEndPoint,
                _clientId,
                Uri.EscapeDataString(RedirectUri),
                _scopes,
                promptType);
        }

        /// <summary>
        /// Base Azure endpoint URI. 
        /// </summary>
        private const string AzureEndPoint = "https://login.microsoftonline.com/{0}/oauth2/v2.0/{1}";

        /// <summary>
        /// Application's client ID.
        /// </summary>
        private readonly string _clientId;

        /// <summary>
        /// Requested permissions scope.
        /// </summary>
        private readonly string _scopes;

        /// <summary>
        /// Microsoft's '/token' endpoint.
        /// </summary>
        private readonly string _tokenEndPoint;

        /// <summary>
        /// Microsoft's '/authorize' endpoint.
        /// </summary>
        private readonly string _authorizeEndPoint;

        /// <summary>
        /// Redeems authorization code for an access token and refresh token via Azure's '/token' endpoint.
        /// </summary>
        /// <param name="authorizationCode">Authorization code that has been obtained during user authentication.</param>
        public async Task RequestAccessTokenAsync(string authorizationCode)
        {
            // construct the request body
            string requestBody = string.Format(
                CultureInfo.InvariantCulture,
                "grant_type=authorization_code&code={0}&scope={1}&redirect_uri={2}&client_id={3}",
                authorizationCode,
                _scopes,
                Uri.EscapeDataString(RedirectUri),
                _clientId);

            // send it via POST request and receive response
            string responseJson = await HttpPostAsync(_tokenEndPoint, requestBody);

            // deserialize JSON response
            var response = DeserializeJson(responseJson);

            // get OAuth 2.0 access token
            AccessToken = GetStringValue(response, "access_token");

            // get OAuth 2.0 refresh token - only provided with 'offline_access' scope
            if (_scopes.Contains("offline_access"))
            {
                RefreshToken = GetStringValue(response, "refresh_token");
            }

            // id_token - JSON Web Token (JWT) - only provided with 'openid' scope
            if (_scopes.Contains("openid"))
            {
                // deserialize id_token
                string[] idTokenEncodedParts = GetStringValue(response, "id_token").Split('.');
                if (idTokenEncodedParts.Length < 2)
                {
                    throw new OAuthAzureException("Unexpected JWT token format.");
                }
                byte[] idTokenData = DecodeFromUrlBase64String(idTokenEncodedParts[1]);
                string idTokenJson = Encoding.UTF8.GetString(idTokenData);
                var idToken = DeserializeJson(idTokenJson);

                // get user name from id_token - only provided with 'profile' scope
                if (_scopes.Contains("profile"))
                {
                    UserName = GetStringValue(idToken, "preferred_username");
                    FullName = GetStringValue(idToken, "name");
                }

                // get user name from id_token - only provided with 'email' scope
                if (_scopes.Contains("email"))
                {
                    Email = GetStringValue(idToken, "email");
                }
            }
        }

        /// <summary>
        /// Preform HTTP POST request to obtain a new OAuth access token. Updates the <see cref="AccessToken"/> property.
        /// </summary>
        public async Task RefreshAccessTokenAsync()
        {
            // construct the request body
            string requestBody = "grant_type=refresh_token&client_id=" + _clientId + "&refresh_token=" + RefreshToken;

            // send it via POST request and receive response
            string responseJson = await HttpPostAsync(_tokenEndPoint, requestBody);

            // deserialize JSON response
            var response = DeserializeJson(responseJson);

            // update OAuth 2.0 access token and refresh token
            AccessToken = GetStringValue(response, "access_token");
            RefreshToken = GetStringValue(response, "refresh_token");
        }

        /// <summary>
        /// Sends data using HTTP POST to the specified URI and receive the JSON response.
        /// </summary>
        /// <param name="requestBody"></param>
        /// <returns></returns>
        private static async Task<string> HttpPostAsync(string uri, string requestBody)
        {
            // convert request to a byte array
            byte[] postBytes = Encoding.UTF8.GetBytes(requestBody);

            // create and post web request
            var request = WebRequest.Create(uri);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBytes.Length;

            // send request data
            using (Stream postStream = request.GetRequestStream())
            {
                await postStream.WriteAsync(postBytes, 0, postBytes.Length);
            }

            // get response
            WebResponse response;
            try
            {
                response = await request.GetResponseAsync();
            }
            catch (WebException error)
            {
                if (error.Status != WebExceptionStatus.ProtocolError)
                {
                    throw;
                }

                response = error.Response;
                if (!response.ContentType.StartsWith("application/json", StringComparison.Ordinal))
                {
                    throw;
                }

                // parse JSON error response and throw a relevant exception
                string responseJson = await ReadResponseToEndAsync(response);
                var errorInfo = DeserializeJson(responseJson);
                string errorCode = GetStringValue(errorInfo, "error");
                string description = GetStringValue(errorInfo, "error_description");
                if (description == null)
                {
                    description = string.Concat("Error '", errorCode, "'.");
                }

                throw new OAuthAzureException(description.ToString(), errorCode);
            }

            return await ReadResponseToEndAsync(response);
        }

        /// <summary>
        /// Reads WebResponse content.
        /// </summary>
        /// <param name="response">WebResponse.</param>
        /// <returns>WebResponse content.</returns>
        private static async Task<string> ReadResponseToEndAsync(WebResponse response)
        {
            try
            {
                using (var reader = new StreamReader(response.GetResponseStream()))
                {
                    // read and return the response body
                    return await reader.ReadToEndAsync();
                }
            }
            finally
            {
                response.Close();
            }
        }

        /// <summary>
        /// Retrieves a string value with the specified name from a dictionary, or null if not present.
        /// </summary>
        /// <param name="dictionary">Dictionary of key/object.</param>
        /// <param name="key">Key.</param>
        /// <returns>String value or null.</returns>
        private static string GetStringValue(Dictionary<string, object> dictionary, string key)
        {
            dictionary.TryGetValue(key, out object value);
            return value != null ? Convert.ToString(value) : null;
        }

        /// <summary>
        /// Decodes a string that encodes binary data in 'base64url' form (https://tools.ietf.org/html/rfc4648#section-5).
        /// </summary>
        /// <param name="base64url">Data encoded in 'base64url' form.</param>
        /// <returns>Decoded data.</returns>
        private static byte[] DecodeFromUrlBase64String(string base64url)
        {
            const char Base64PadCharacter = '=';
            const string Base64Character62 = "+";
            const string Base64Character63 = "/";
            const string Base64UriCharacter62 = "-";
            const string Base64UriCharacter63 = "_";

            string base64 = base64url.Replace(Base64UriCharacter62, Base64Character62);
            base64 = base64.Replace(Base64UriCharacter63, Base64Character63);
            base64 = base64.PadRight((base64url.Length + 3) & ~3, Base64PadCharacter);
            return Convert.FromBase64String(base64);
        }

        /// <summary>
        /// Deserializes the supplied JSON string.
        /// </summary>
        /// <param name="json">JSON string to deserialize.</param>
        /// <returns>Deserialized JSON tree.</returns>
        private static Dictionary<string, object> DeserializeJson(string json)
        {
            // Microsoft's JSON serializer (https://www.nuget.org/packages/System.Text.Json/)
            // (available for .NET Framework 4.7.2 or higher)
            return System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(json);

            // Newtonsoft.Json (https://www.nuget.org/packages/Newtonsoft.Json/)
            // (available for .NET Framework 2.0 or higher
            //return Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
        }
    }
}
