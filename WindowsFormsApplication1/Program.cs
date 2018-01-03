using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            SplashForm splash = new SplashForm();
            DateTime end = DateTime.Now + TimeSpan.FromSeconds(5);
            splash.Show();
            while (end > DateTime.Now)
            {
                Application.DoEvents();
            }
            splash.Close();
            splash.Dispose();
            Application.Run(new Form1());
        }
    }
}
