using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
namespace SillyUser
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(Application_UnhandledException);

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var f = new Form1();
                Application.Run(f);
            }catch (Exception ex)
            {
                ShowErrorBox(ex);
            }
            
        }

        private static void Application_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ShowErrorBox((System.Exception)e.ExceptionObject); 
        }

        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            
            ShowErrorBox(e.Exception);
        }
        private static void ShowErrorBox(Exception ex)
        {
            var errorMessage = ex.ToString() + "\r\n" + new StackTrace(ex).ToString();

            MessageBox.Show(errorMessage, "Unrecoverable exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
