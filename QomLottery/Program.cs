using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QomLottery
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
           // IronXL.License.LicenseKey = IronXLKeygen.GenerateKey();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
           // Application.Run(new ShowLottery("شهرداری قم "+"\n"+"برنده خوش شاش"+ "\n" + "منشسیتبمسنیت س"+ "\n" + "یبلایبلایبلایبلایبلا"+ "\n" + "یبلایبلایبلایبلابیل"));
        }
    }
}
