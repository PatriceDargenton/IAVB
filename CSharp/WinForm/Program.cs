
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IAVB
{
    static class Program
    {

#if DEBUG
        public static bool bDebug = true;
        //public static bool bRelease = false;
#else
        public static bool bDebug = false;
        //public static bool bRelease = true;
#endif

        public const string sDateVersionAppli = "18/06/2017";
        public const string sVersionAppli = "3.12";
        // Pas de My.Application en C# :
        //public readonly sVersionAppli = My.Application.Info.Version.Major + "." 
        //    + My.Application.Info.Version.Minor + My.Application.Info.Version.Build;

        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmIAVB());
        }
    }
}
