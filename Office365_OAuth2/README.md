Office365_OAuth2
================

**ImapOAuthWpfApp**, **GraphAuthWpfApp**, **EwsOAuthWpfApp**: Three WPF applications and a helper library that show how
to authenticate to Office 365 (Microsoft 365 / Exchange Online) using OAuth 2.0
in *delegated* mode (suitable for user-attended apps) and retrieve
a list of recent messages with [Rebex Secure Mail](https://www.rebex.net/secure-mail.net/)
using IMAP, MS Graph or EWS (Exchange Web Services) protocols. Targets .NET Framework 4.7.2 and .NET 6.0.

For details, see the following articles:
- [How to use OAuth2.0 authentication for Office 365 with Rebex Secure Mail and IMAP/EWS/POP3/SMTP](https://blog.rebex.net/oauth2-office365-rebex-mail)
- [How to use OAuth2.0 authentication for Office 365 with Rebex Secure Mail and MS Graph API](https://blog.rebex.net/office365-graph-oauth-delegated)
- [How to register your app for Office 365 with OAuth 2.0 authentication](https://blog.rebex.net/registering-app-for-oauth2-office365)

For a backport of this sample to .NET Framework 3.5 SP1, see [Office365_OAuth2_Legacy](../Office365_OAuth2_Legacy).

There is also the [Office365_OAuth2_IdentityClient](../Office365_OAuth2_IdentityClient) variant that uses [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) package for the OAuth 2.0 authentication flow.

![Screenshot](https://raw.githubusercontent.com/rebexnet/RebexExtras/master/Office365_OAuth2/screenshot.png)
