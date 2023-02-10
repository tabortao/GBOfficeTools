using System;
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

    //internal static class Program
    //{
    //    private const int WS_SHOWNORMAL = 1;

    //    [DllImport("User32.dll")]
    //    private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);

    //    [DllImport("User32.dll")]
    //    private static extern bool SetForegroundWindow(IntPtr hWnd);

    //    /// <summary>
    //    /// 应用程序的主入口点。
    //    /// </summary>
    //    [STAThread]
    //    private static void Main()
    //    {
    //        Process instance = GetRunningInstance();
    //        if (instance == null)
    //        {
    //            Application.EnableVisualStyles();
    //            Application.SetCompatibleTextRenderingDefault(false);
    //            Application.Run(new MainForm());//在这启动主窗体。
    //        }
    //        else
    //        {
    //            HandleRunningInstance(instance);
    //        }
    //    }

    //    /// <summary>
    //    /// 获取当前是否具有相同进程。
    //    /// </summary>
    //    /// <returns></returns>
    //    public static Process GetRunningInstance()
    //    {
    //        Process current = Process.GetCurrentProcess();
    //        Process[] processes = Process.GetProcessesByName(current.ProcessName);
    //        //遍历正在有相同名字运行的例程
    //        foreach (Process process in processes)
    //        {
    //            //忽略现有的例程
    //            if (process.Id != current.Id)
    //                //确保例程从EXE文件运行
    //                if (System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("/", "\\") == current.MainModule.FileName)
    //                    return process;
    //        }
    //        return null;
    //    }

    //    /// <summary>
    //    /// 激活原有的进程。
    //    /// </summary>
    //    /// <param name="instance"></param>
    //    public static void HandleRunningInstance(Process instance)
    //    {
    //        ShowWindowAsync(instance.MainWindowHandle, WS_SHOWNORMAL);
    //        SetForegroundWindow(instance.MainWindowHandle);
    //    }
    //}
}