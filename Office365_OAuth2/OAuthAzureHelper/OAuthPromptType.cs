using System;

namespace Rebex.Samples
{
    /// <summary>
    /// Defines possible OAuth prompt types.
    /// See https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code for details.
    /// </summary>
    public class OAuthPromptType
    {
        /// <summary>
        /// Don't request credentials if already signed-on.
        /// </summary>
        public const string Default = "";

        /// <summary>
        /// Always request credentials.
        /// </summary>
        public const string Login = "login";

        /// <summary>
        /// Never request credential.
        /// If the request can't be completed silently using single sign-on, Microsoft identity platform will return an "interaction_required" error.
        /// </summary>
        public const string None = "none";

        /// <summary>
        /// Ask the user to grant permissions to the app.
        /// </summary>
        public const string Consent = "consent";

        /// <summary>
        /// Ask the user to select the account to use.
        /// </summary>
        public const string SelectAccount = "select_account";
    }
}
