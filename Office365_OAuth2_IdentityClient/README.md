Office365_OAuth2_IdentityClient
===============================

## App-only (unattended) authentication (for services and deamons)

**EwsOAuthAppOnlyConsole_IdentityClient**   
**ImapOAuthAppOnlyConsole_IdentityClient**   
**Pop3OAuthAppOnlyConsole_IdentityClient**

These console applications show how to authenticate to Microsoft 365 (Office 365, Exchange Online)
with OAuth 2.0 using [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) 
in *app-only* mode (suitable for unattended services/deamons)
and retrieve a list of recent messages with [Rebex Secure Mail](https://www.rebex.net/secure-mail.net/)
using EWS (Exchange Web Services), IMAP, and POP3 protocols. Targets .NET Framework 4.6 and .NET 6.0.

Step-by-step guides that describe how to configure this in Microsoft 365:

- [Office 365 and EWS with OAuth 2.0 authentication in unattended (app-only) mode](https://blog.rebex.net/office365-ews-oauth-unattended)
- [Office 365 and IMAP or POP3 with OAuth 2.0 authentication in unattended (app-only) mode](https://blog.rebex.net/office365-imap-pop3-oauth-unattended)

-----

## Delegated authentication (for user apps)

**EwsOAuthWpfApp_IdentityClient**   
**ImapOAuthWpfApp_IdentityClient**   
**SmtpOAuthWpfApp_IdentityClient**

These WPF applications show how to authenticate to Microsoft 365 (Office 365, Exchange Online)
with OAuth 2.0 using [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) 
in *delegated* mode (suitable for user-attended apps)
and retrieve a list of recent messages with [Rebex Secure Mail](https://www.rebex.net/secure-mail.net/)
using EWS (Exchange Web Services), IMAP, and SMTP protocols. Targets .NET Framework 4.6 and .NET 6.0.

For a version for .NETFramework 4.7.2 and .NET 6.0 that does not use Microsoft.Identity.Client,
see [Office365_OAuth2](../Office365_OAuth2). That version makes it easier to understand how
OAuth 2.0 flow actually works under the hood. For details, see the following articles:

- [How to register your app for Office 365 with OAuth 2.0 authentication](https://blogtest.rebex.net/registering-app-for-oauth2-office365)
- [How to use OAuth2.0 authentication for Office 365 with Rebex Secure Mail](https://blog.rebex.net/oauth2-office365-rebex-mail)

![Screenshot](https://raw.githubusercontent.com/rebexnet/RebexExtras/master/Office365_OAuth2/screenshot.png)
