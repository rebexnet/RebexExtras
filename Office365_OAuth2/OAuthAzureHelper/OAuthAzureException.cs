using System;

namespace Rebex.Samples
{
    /// <summary>
    /// Represents an exception that occured during OAuth2 authentication.
    /// </summary>
    public class OAuthAzureException : Exception
    {
        public string ErrorCode { get; private set; }

        public OAuthAzureException(string message) : this(message, (Exception)null)
        {
        }

        public OAuthAzureException(string message, string errorCode) : base(message)
        {
            ErrorCode = errorCode;
        }

        public OAuthAzureException(string message, Exception innerException) : base(message, innerException)
        {
            ErrorCode = "unspecified";
        }
    }
}
