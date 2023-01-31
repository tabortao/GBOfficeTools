using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace BookmarksTool
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        private static void Main()
        {
            // 小文——在C# WinForm中如何使当前应用程序只允许启动一个实例 https://www.cnblogs.com/jaxu/archive/2009/05/05/1450083.html
            bool createNew;
            using (System.Threading.Mutex m = new System.Threading.Mutex(true, Application.ProductName, out createNew))
            {
                if (createNew)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new MainForm());
                }
                else
                {
                    MessageBox.Show("GBOfficeTools已经打开，请在系统托盘中查看！");
                }
            }
        }
    }
}
