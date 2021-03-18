Office365_OAuth2
================

Two WPF applications and a helper library that show how
to authenticate to Office365 (Exchange Online) using OAuth 2.0 and retrieve
a list of recent messages with [Rebex Secure Mail](https://www.rebex.net/secure-mail.net/)
using IMAP or EWS (Exchange Web Services) protocols. Targets .NET Framework 4.7.2 and .NET 5.0.

See the blog post at https://blog.rebex.net/oauth2-office365-rebex-mail for details on how it works.

For a backport of this sample to .NET Framework 3.5 SP1, see [Office365_OAuth2_Legacy](../Office365_OAuth2_Legacy).

There is also the [Office365_OAuth2_IdentityClient](../Office365_OAuth2_IdentityClient) variant that uses [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/) package for the OAuth 2.0 authentication flow.

![Screenshot](https://raw.githubusercontent.com/rebexnet/RebexExtras/master/Office365_OAuth2/screenshot.png)
