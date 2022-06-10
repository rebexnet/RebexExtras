using System;

namespace Rebex.Samples
{
    internal static class LicenseKey
    {
        /// <summary>
        /// If you have already purchased a Rebex component license at https://www.rebex.net/,
        /// put your Rebex license key below. See https://www.rebex.net/kb/license-keys/ for
        /// instructions on getting your key.
        /// Otherwise, to start your 30-day evaluation period, get your trial key
        /// from https://www.rebex.net/support/trial/
        /// </summary>
        public static string Value
        {
            get
            {
                string key = null; // put your license key here

                // if no key is set, try reading it from REBEX_KEY environment variable.
                return key ?? Environment.GetEnvironmentVariable("REBEX_KEY");
            }
        }
    }
}
