using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SoundCheckTCPClient
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
           //registerForm = new Login();

            //DllDebug registerForm = new DllDebug();

            //registerForm.ShowDialog();
           // if (registerForm.DialogResult == DialogResult.OK)
         //   {
                try
                {
                    Application.Run(new Project());

                    //MainSelect form = new MainSelect();
                    //form.ShowDialog();
                    //if (form.DialogResult == DialogResult.OK)
                    //{
                    //    Application.Run(new Main());
                    //}
                }
                catch { }
            //}
            //else
            //{
            //    return;
            }
        }
}
