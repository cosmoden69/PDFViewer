using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFViewer
{
    static class Program
    {
        public static string G_UP_Path = @"C:\haesungHASP\";
        public static string G_WD_Path = @"C:\haesungHASP\L_Image\";
        public static string G_Create_ = @"C:\haesungHASP\CrtWord\";
        public static int GUserSize = 70;

        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args == null || args.Length < 1)
            {
                MessageBox.Show("파라미터가 없습니다");
                return;
            }

            string docFile = args[0];
            if (!File.Exists(docFile))
            {
                MessageBox.Show("파일이 없습니다");
                return;
            }

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1(args));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
