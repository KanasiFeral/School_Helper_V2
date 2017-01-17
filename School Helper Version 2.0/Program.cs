using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using School_Helper_Version_2._0.Forms;
using System.Data.OleDb;

namespace School_Helper_Version_2._0
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ConnectorAccess ClassConSQL = new ConnectorAccess();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Splashscreen(ClassConSQL));
        }
    }
}
