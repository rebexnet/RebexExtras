Office365_OAuth2_IdentityClient
===============================

Two WPF applications that show how to authenticate to Office365 (Exchange Online)
with OAuth 2.0 using [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) 
and retrieve a list of recent messages with [Rebex Secure Mail](https://www.rebex.net/secure-mail.net/)
using IMAP or EWS (Exchange Web Services) protocols. Targets .NET Framework 4.6 and .NET 5.0.

For a version for .NETFramework 4.7.2 and .NET 5.0 that does not use Microsoft.Identity.Client,
see [Office365_OAuth2](../Office365_OAuth2). That version makes it easier to understand how
OAuth 2.0 flow actually works under the hood. See the blog post at
https://blog.rebex.net/oauth2-office365-rebex-mail for details.

![Screenshot](https://raw.githubusercontent.com/rebexnet/RebexExtras/master/Office365_OAuth2/screenshot.png)
