using AutoUpdaterDotNET;
using System;
using System.Windows.Threading;

namespace BookmarksTool.LeiTools
{
    internal class AutoUpdate
    {
        public static void CheckUpdate()
        {
            //Todo：实现软件更新功能 20230116
            // https://gitee.com/yhuse/SunnyUI.AutoUpdater.NET
            // https://github.com/ravibpatel/AutoUpdater.NET ，感谢开源软件“AutoUpdater.NET”，网页翻译为汉语学习。

            // 1.编制XML文件

            // 2.设置AutoUpdaterDotNET
            //您可以通过添加以下代码来打开错误报告。如果这样做 AutoUpdater.NET 将显示错误消息，
            //如果没有可用的更新或无法从 Web 服务器访问 XML 文件。
            AutoUpdater.ReportErrors = true;

            //如果您不想下载最新版本的应用程序，而只想打开XML文件的url标签之间的URL，则需要添加以下行和上面的代码。
            //AutoUpdater.OpenDownloadPage = true;
            //打开 https://gitee.com/yaoleistable/GBOfficeToolsReleases/releases/tag/UpdateXML 更新Update.xml。
            //代替自建HTTP服务器，路径不要改变、文件名不要改名
            // http://192.168.31.184/BookmarksToolUpdate/Update.xml
            //如果您不希望用户在按下更新对话框的“稍后提醒”按钮时选择“稍后提醒”时间，则需要使用上述代码添加以下行。
            //AutoUpdater.LetUserSelectRemindLater = false;
            //AutoUpdater.RemindLaterTimeSpan = RemindLaterFormat.Days;
            //AutoUpdater.RemindLaterAt = 2;
            // 强制更新
            //可以通过将“必需属性”设置为 true 并将“更新模式”设置为值 或 来启用强制更新。 选项将隐藏标准更新对话框中的稍后提醒，跳过和关闭按钮。 
            //选项将跳过标准更新对话框，无需用户交互即可开始下载和更新应用程序。 选项还将忽略 OpenDownloadPage 标志的值。
            //AutoUpdater.Mandatory = true;
            //AutoUpdater.UpdateMode = Mode.Forced;

            //指定更新窗体的大小
            //AutoUpdater.UpdateFormSize = new System.Drawing.Size(900, 600);

            //经常检查更新:可以在计时器中调用 Start 方法来频繁检查更新。
            //DispatcherTimer timer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
            //timer.Tick += delegate
            //{
            //    AutoUpdater.Start("https://gitee.com/yaoleistable/GBOfficeToolsReleases/releases/download/UpdateXML/Update.xml");
            //};
            //timer.Start();
            // 3.启动更新
            AutoUpdater.Start("https://gitee.com/yaoleistable/GBOfficeToolsReleases/releases/download/UpdateXML/Update.xml");
            //System.Windows.Forms.MessageBox.Show("软件更新完成");                    
            //string updates = @"https://gitee.com/yaoleistable/GBOfficeTool/blob/master/Updates.md";
            //System.Diagnostics.Process.Start(updates);

        }

        public static void AutoCheckUpdate()
        {
            //软件启动，自动检查更新，强制升级。
            AutoUpdater.Mandatory = true;
            AutoUpdater.UpdateMode = Mode.Forced;
            AutoUpdater.Start("https://gitee.com/yaoleistable/GBOfficeToolsReleases/releases/download/UpdateXML/Update.xml");
        }
    }
}