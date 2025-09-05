using System;
using System.Windows.Forms;

namespace SmtpOAuthWinFormApp
{
    public static class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Control.CheckForIllegalCrossThreadCalls = true;

            Application.Run(new MainForm());
        }
    }
}
