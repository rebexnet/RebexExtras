Rebex Extras
============

Collection of additional useful sample apps and libraries written using
[Rebex components](https://www.rebex.net/total-pack/).


## WPF apps using Rebex Secure Mail to access Office365 with OAuth 2.0 authentication

[Office365_OAuth2](Office365_OAuth2) - A set of two WPF applications (and a helper library) that show how
to authenticate to Office365 (Exchange Online) using OAuth 2.0 and retrieve
a list of recent messages with [Rebex Secure Mail](https://www.rebex.net/secure-mail.net/)
using IMAP, MS Graph or EWS (Exchange Web Services) protocols. Targets .NET Framework 4.7.2 and .NET 6.0.

[Office365_OAuth2_Legacy](Office365_OAuth2_Legacy) - A backport of the previous sample app to
.NET Framework 3.5 SP1. Uses [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json/) for
parsing OAuth responses

[Office365_OAuth2_IdentityClient](Office365_OAuth2_IdentityClient) - Another variant
of Office365_OAuth2. This one uses [Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/)
package for the OAuth 2.0 authentication flow. Also includes a third app that shows how to
authenticate to EWS using *app-only* mode (suitable for unattended services/deamons).


## Usage

All samples reference Rebex assemblies from [NuGet.org](https://www.nuget.org/profiles/rebex).
In order to start using them, get your [30-day trial licence key](https://www.rebex.net/support/trial-key.aspx)
and put it into [LicenseKey.cs](LicenseKey.cs) file. If you have already purchased a license,
use a [full key](https://www.rebex.net/kb/license-keys/) instead.


## Licensing

All the samples are available under freeware [Rebex Sample Code License](LICENSE.txt).

For information on licensing Rebex components, see [Licensing FAQ](https://www.rebex.net/shop/faq/) for details.
