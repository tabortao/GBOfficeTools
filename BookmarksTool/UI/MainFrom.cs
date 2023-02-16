using Microsoft.Win32;
using Sunny.UI;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace BookmarksTool
{
    public partial class MainForm : UIForm

    {
        public static MainForm form1; //定义为静态变量，并在构造方法中引用Form1，可以实现在另外的一个类中调用窗体的控件或方法

        public string help =
                "    一、Word书签工具使用帮助：\r\n" +
                "       （1）选择Excel文件，选择要转换的Words文件，点击“批量生成Word”按钮。\r\n" +
                "    二、WordPDF使用帮助：\r\n" +
                "       （1）方式一：选择要转换的Words文件，点击“Word批量转PDF”按钮。\r\n" +
                "       （2）方式二：选择文件夹路径（存放Word的文件夹）,点击“Word批量转PDF”按钮\r\n" +
                "    三、关于软件：\r\n" +
                "       （1）作者：可持续学园。\r\n" +
                "       （2）《GBOfficeTools软件说明与帮助》，点击下方“软件说明”按钮打开浏览器进行阅读。 \r\n" +
                "       （3）关注公众号“可持续学园”，回复“GBOfficeTools”获取软件下载地址，留言提出宝贵建议。\r\n";

        public MainForm()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            form1 = this; //构造方法中引用Form1
            this.rtxt_help.Text = help;
            this.uiStyleManager1.DPIScale = true; // 防止缩放产生问题 vV2.0.8条修复问题
            string version = Application.ProductVersion.ToString();
            this.lab_Version.Text = "当前版本：" + version.Substring(0, 5) + "\n";

            //Excel模板书签设置 读取AppConfig.ini
            this.txt_SheetPageNo.Text = LeiTools.ConfigHelper.IniHelper.ReadString("Excel模板书签设置", "书签所在Sheet页", "NA");
            this.txt_bookmarkNameNo.Text = LeiTools.ConfigHelper.IniHelper.ReadString("Excel模板书签设置", "书签名所在列", "NA");
            this.txt_bookmarkValueNo.Text = LeiTools.ConfigHelper.IniHelper.ReadString("Excel模板书签设置", "书签值所在列", "NA");
            this.txt_bookmarkNo.Text = LeiTools.ConfigHelper.IniHelper.ReadString("Excel模板书签设置", "书签个数", "NA");
        }

        #region 自动更新

        private void btn_Update_Click(object sender, EventArgs e)
        {
            //调用自动更新函数，进行软件更新
            LeiTools.AutoUpdate.CheckUpdate();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 读取系统设置，APP是否开机启动
            ck_Selfstart.Checked = Convert.ToBoolean(LeiTools.ConfigHelper.IniHelper.ReadString("App设置", "IsAutoBoot", "NA"));
            //读取系统设置，截图 OCR快捷键
            txt_Screencapture.Text = LeiTools.ConfigHelper.IniHelper.ReadString("App设置", "Screencapture", "NA");


            //ck_Selfstart.Checked = LeiTools.ConfigHelper.JSONHelper.ReadJSON<bool>("IsAutoBoot");
            //打开软件后，自动检查是否有可用更新,如有更新，弹出对话框询问是否更新，点击更新后，自动更新。
            LeiTools.AutoUpdate.AutoCheckUpdate();

            //启动后自动最小化
            //this.WindowState = FormWindowState.Minimized;

            // 注册热键
            //注册热键Shift+S，Id号为100。HotKey.KeyModifiers.Shift也可以直接使用数字4来表示。   
            //LeiTools.SystemOp.HotKey.RegisterHotKey(Handle, 100, LeiTools.SystemOp.HotKey.KeyModifiers.Shift, Keys.S);
            //注册热键Ctrl+B，Id号为101。HotKey.KeyModifiers.Ctrl也可以直接使用数字2来表示。   
            //LeiTools.SystemOp.HotKey.RegisterHotKey(Handle, 101, LeiTools.SystemOp.HotKey.KeyModifiers.Ctrl, Keys.B);
            //注册热键Ctrl+Alt+D，Id号为102。HotKey.KeyModifiers.Alt也可以直接使用数字1来表示。   
            //LeiTools.SystemOp.HotKey.RegisterHotKey(Handle, 102, LeiTools.SystemOp.HotKey.KeyModifiers.Alt | LeiTools.SystemOp.HotKey.KeyModifiers.Ctrl, Keys.D);
            //注册热键F4，Id号为103。   
            LeiTools.SystemOp.HotKey.RegisterHotKey(Handle, 103, LeiTools.SystemOp.HotKey.KeyModifiers.None, Keys.F4);
        }
        private void btn_ShortcutKey_Click(object sender, EventArgs e)
        {
            //自定义 修改截图ORC快捷键
            //LeiTools.ConfigHelper.IniHelper.WriteString("App设置", "Screencapture", Convert.ToString(this.txt_Screencapture.Text));
            //MessageBox.Show("截图 OCR快捷键设置成为了"+txt_Screencapture.Text);
            MessageBox.Show("截图 OCR快捷键默认F4，暂不允许修改");
        }
        /// <summary>
        /// 注册热键
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;
            //按快捷键    
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    switch (m.WParam.ToInt32())
                    {
                        case 100:    //按下的是Shift+S   
                            //MessageBox.Show("按下的是Shift+S");
                            break;
                        case 101:    //按下的是Ctrl+B   
                                     //此处填写快捷键响应代码   
                            //MessageBox.Show("按下的是Ctrl+B");
                            break;
                        case 102:    //按下的是Alt+D   
                                     //此处填写快捷键响应代码   
                            //MessageBox.Show("按下的是Ctrl+Alt+D");
                            break;
                        case 103:
                            //this.Text = "F4";
                            string screencapture = System.Environment.CurrentDirectory + @"\Screencapture\screencapture.exe";
                            Process.Start(screencapture);
                            //MessageBox.Show("按下的是F4");
                            break;
                    }
                    break;
            }
            base.WndProc(ref m);
        }

        #endregion 自动更新

        #region Word书签工具

        public static string excelFilePath; // 定义静态变量 excelFilePath，用于存储excel模板文件路径
        public static string[] wordsPath;// 定义静态变量 wordsPath，用于存储选中的多个word模板文件路径

        public void TextBoxMsg(string msg)  //输出消息到多行文本框
        {
            textBox1.AppendText(msg + "\r\n");
            //textBox1.Text += msg + "\r\n";
        }

        private void Btn_start_Click(object sender, EventArgs e)
        {
            bool isDateExpired = LeiTools.Keygen.DateTimeHelper.IsDateExpired(2025); // 设置软件过期时间为2025年
            if (isDateExpired)
            {
                #region 执行Word报告生成

                //Console.WriteLine(@"正在运行……");
                //textBox1.Text += "正在运行……";
                //string mCode = RegInfo.GetMachineCode();
                //string leiCode = "681348742402123950127693";
                //textBox1.ResetText();
                textBox1.Clear();
                //textBox1.Text = string.Empty;
                //textBox1.Text = ""; //清空文本框内容，为下次执行扫除干净
                //System.Threading.Thread.Sleep(10);

                TextBoxMsg("批量生成Word报告正在运行，请稍后……");
                var sw = new Stopwatch();
                sw.Start();  //开始计时
                if (txt_Excel1.Text.Length != 0 && txt_Words1.Text.Length != 0)
                {
                    //普通循环运算，进行word书签替换
                    LeiTools.AsposeOffice.BookmarksReplace.ReportMaker(excelFilePath, wordsPath);
                    //并行运算，进行word书签替换
                    //LeiTools.AsposeOffice.BookmarksReplace.ParallelReportMaker(excelFilePath, wordsPath);
                }
                else if (txt_Excel1.Text.Length == 0 && txt_Words1.Text.Length == 0)
                {
                    //如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
                    //普通循环运算，进行word书签替换
                    LeiTools.AsposeOffice.BookmarksReplace.ReportMaker();
                    //并行运算，进行word书签替换
                    //LeiTools.AsposeOffice.BookmarksReplace.ParallelReportMaker();
                }
                else if (txt_Excel1.Text.Length == 0 || txt_Words1.Text.Length == 0)
                {
                    MessageBox.Show("请选择Excel模板文件或选择Word文件");
                }
                else
                {
                    MessageBox.Show("文件选择有误，请重新选择！");
                }

                //LeiTools.AsposeOffice.BookmarksReplace.StartBookmarksReplace2();
                sw.Stop();
                //Console.WriteLine("运行结束,用时{0}秒！按任意键结束", sw.Elapsed);
                MainForm.form1.TextBoxMsg("运行结束,用时" + sw.Elapsed + "秒！");

                #endregion 执行Word报告生成
            }
            else
            {
                MessageBox.Show("软件授权过期，请关注微信公众号“可持续学园”留言，免费获取更新。");
            }
        }

        private void but_SelectExcel_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            OpenFileDialog dialog = new OpenFileDialog
            {
                //是否支持多个文件的打开？
                Multiselect = false,
                //标题
                Title = "请选择Excel文件",
                //文件类型
                Filter = "Excel文件(*.xlsx)|*.xlsx",//或"图片(*.jpg;*.bmp)|*.jpg;*.bmp"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                excelFilePath = dialog.FileName; //选择的excel文件路径，赋值给静态变量 excelFilePath
                //TextBoxMsg(excelFilePath);
                txt_Excel1.AppendText(excelFilePath);

                //获取选择的文件路径
                //string str = "";
                //for (int i = 0; i < dialog.FileNames.Length; i++) //根据数组长度定义循环次数
                //    str += dialog.FileNames.GetValue(i).ToString();//获取文件文件名
                //TextBoxMsg(str);
                //MessageBox.Show(String.Format("Name={0} !", str));
            }
            //return FilePath;
        }

        private void btn_SelectWords_Click(object sender, EventArgs e)
        {
            textBox1.Clear();

            OpenFileDialog dialog = new OpenFileDialog
            {
                //是否支持多个文件的打开？
                Multiselect = true,
                //标题
                Title = "请选择Word文件",
                //文件类型
                Filter = "Word文件(*.doc;*.docx)|*.doc;*.docx",//或"图片(*.jpg;*.bmp)|*.jpg;*.bmp"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                wordsPath = dialog.FileNames;
                foreach (var wordPath in wordsPath)
                {
                    //TextBoxMsg(wordPath);
                    txt_Words1.AppendText(wordPath + "\r\n");
                }

                //string[] filesPath = dialog.FileNames;
                //foreach (var filePath in filesPath)
                //{
                //    TextBoxMsg(filePath);
                //    txt_Words.AppendText(filePath);
                //}

                //获取选择的文件路径
                //string str = "";
                //for (int i = 0; i < dialog.FileNames.Length; i++) //根据数组长度定义循环次数
                //    str += dialog.FileNames.GetValue(i).ToString();//获取文件文件名
                //TextBoxMsg(str);
                //MessageBox.Show(String.Format("Name={0} !", str));
            }
            //return FilePath;
        }

        /// <summary>
        /// 清除所有文本框内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Clear_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            txt_Excel1.Clear();
            txt_Words1.Clear();
        }

        #endregion Word书签工具

        #region Word转PDF工具

        public static string[] wordsPath2;
        public static string folderPath2;

        public void TextBoxMsg2(string msg)  //输出消息到多行文本框
        {
            textBox2.AppendText(msg + "\r\n");
            //textBox1.Text += msg + "\r\n";
        }

        private void btn_word2pdf_Click(object sender, EventArgs e)
        {
            //textBox1.ResetText();
            textBox2.Clear();
            //textBox1.Text = "";
            //System.Threading.Thread.Sleep(1000);

            //textBox1.Text = ""; //清空文本框内容，为下次执行扫除干净
            TextBoxMsg2("Word批量转PDF正在运行，请稍等……");
            var sw = new Stopwatch();
            sw.Start(); //开始计时
            //var start2 = new LeiTools.MSOffice.Word.Word2PDF();
            //start2.StartWord2PDF();
            if (txt_Words2.Text.Length != 0 || (txt_Folder2.Text.Length != 0 && txt_Words2.Text.Length != 0))
            {
                //如果选择了多个word文件，或者既选择了文件夹，又选择了多个word文件时，将多个word文件转为PDF
                LeiTools.AsposeOffice.WordToPDF.Word2PDF(wordsPath2);
            }
            else if (txt_Folder2.Text.Length != 0)
            {
                //如果选择了文件夹路径，把文件夹内的word转为pdf
                LeiTools.AsposeOffice.WordToPDF.Word2PDF(folderPath2);
            }
            else if (txt_Words2.Text.Length == 0 && txt_Folder2.Text.Length == 0)
            {
                //如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
                LeiTools.AsposeOffice.WordToPDF.Word2PDF();
            }
            else
            {
                MessageBox.Show("选择的Word文件有误，请重试！");
            }
            sw.Stop();//计时结束
            MainForm.form1.TextBoxMsg2("Word批量转PDF已完成，用时" + sw.Elapsed + "秒");
        }

        private void btn_SelectWords2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            OpenFileDialog dialog = new OpenFileDialog
            {
                //是否支持多个文件的打开？
                Multiselect = true,
                //标题
                Title = "请选择Word文件",
                //文件类型
                Filter = "Word文件(*.doc;*.docx)|*.doc;*.docx",//或"图片(*.jpg;*.bmp)|*.jpg;*.bmp"
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                wordsPath2 = dialog.FileNames;
                foreach (var wordPath in wordsPath2)
                {
                    //TextBoxMsg(wordPath);
                    txt_Words2.AppendText(wordPath + "\r\n");
                }
            }
        }

        private void btn_SelectFolder2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            //FolderBrowserDialog dialog = new FolderBrowserDialog();
            //dialog.Description = "请选择文件夹";

            //if (dialog.ShowDialog() == DialogResult.OK)
            //{
            //    folderPath2 = dialog.SelectedPath;
            //    //TextBoxMsg(folderPath);
            //    txt_Folder2.AppendText(folderPath2);
            //}

            //使用Ookii 优化文件夹选择对话框
            var folderDialog = new Ookii.Dialogs.WinForms.VistaFolderBrowserDialog
            {
                Description = "选择文件夹"
            };

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                folderPath2 = folderDialog.SelectedPath;
                txt_Folder2.AppendText(folderPath2);
            }
        }

        private void btn_Clear2_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            txt_Words2.Clear();
            txt_Folder2.Clear();
        }

        #endregion Word转PDF工具

        private void btn_OpenHelp_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://mp.weixin.qq.com/s/wy68Un1If0Esz0nrYunSbw");
        }

        /// <summary>
        /// 修改Excel模板书签设置设置值，修改AppConfig.ini
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_ChangeSetting_Click(object sender, EventArgs e)
        {
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签所在Sheet页", this.txt_SheetPageNo.Text);
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签名所在列", this.txt_bookmarkNameNo.Text);
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签值所在列", this.txt_bookmarkValueNo.Text);
        }

        private void txt_SheetPageNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //只允许输入数字

            if (e.KeyChar != '\b')//这是允许输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//这是允许输入0-9数字
                {
                    e.Handled = true;
                }
            }
        }

        private void txt_bookmarkNameNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //只允许输入数字

            if (e.KeyChar != '\b')//这是允许输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//这是允许输入0-9数字
                {
                    e.Handled = true;
                }
            }
        }

        private void txt_bookmarkValueNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            //只允许输入数字

            if (e.KeyChar != '\b')//这是允许输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//这是允许输入0-9数字
                {
                    e.Handled = true;
                }
            }
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签所在Sheet页", "0");
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签名所在列", "3");
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签值所在列", "2");
            LeiTools.ConfigHelper.IniHelper.WriteString("Excel模板书签设置", "书签个数", "500");
            txt_SheetPageNo.Text = "0";
            txt_bookmarkNameNo.Text = "3";
            txt_bookmarkValueNo.Text = "2";
            txt_bookmarkNo.Text = "500";
        }

        /// <summary>
        /// 生成选取的文件的名称，不含拓展名,存放于软件目录的\Data\获取的文件名称.txt
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_FileNames_Click(object sender, EventArgs e)
        {
            string fileNamePath = System.IO.Directory.GetCurrentDirectory() + @"\Data\获取的文件名称.txt";
            OpenFileDialog dialog = new OpenFileDialog
            {
                //是否支持多个文件的打开？
                Multiselect = true,
                //标题
                Title = "请选择文件名",
                //文件类型
                Filter = "所有文件(*.*)|*.*",
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string[] fileNamePaths = dialog.FileNames;
                if (File.Exists(fileNamePath))
                {
                    File.Delete(fileNamePath);
                }
                StreamWriter sw = new StreamWriter(fileNamePath);
                foreach (var f in fileNamePaths)
                {
                    try
                    {
                        sw.WriteLine(Path.GetFileNameWithoutExtension(f));
                    }
                    catch (Exception er)
                    {
                        MessageBox.Show("Exception: " + er.Message);
                    }
                }
                sw.Close();
                //打开记事本程序
                System.Diagnostics.Process.Start("NOTEPAD.exe", fileNamePath);
            }
        }

        // 20230210：修改本方法，更加实用方便
        private void btn_ProjectFolder_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog dialog = new FolderBrowserDialog();
            //dialog.Description = "请选择文件夹";
            //优化文件夹选择对话框
            var dialog = new Ookii.Dialogs.WinForms.VistaFolderBrowserDialog
            {
                Description = "选择文件夹"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var folderPath3 = dialog.SelectedPath;
                LeiTools.IOHelper.BatchCreateFolder(folderPath3);
                MessageBox.Show("生成项目文件夹成功！");
            }
        }

        // 20230210：取消本方法，改用BatchCreateFolder(string folderPath)
        public string CreatProjectPath(string folderPath, string path)
        {
            string iniPath = Directory.GetCurrentDirectory() + @"\Data\ProjectFolder.ini";
            string projectName = @"\P23-XX XXX项目";
            //string path1 = folderPath + projectName + @"\00 合同";
            return folderPath + projectName + @"\" + LeiTools.ConfigHelper.IniHelper.ReadString(iniPath, "项目文件夹", path, "NA");
        }

        private void contextMain(object sender, MouseEventArgs e)
        {
        }

        //调整窗体大小的时候触发的事件
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)//当窗体设置值为最小化时
            {
                notifyIcon1.Visible = true;//该控件可见
                this.ShowInTaskbar = false;//在任务栏中显示该窗口
            }
            else
            {
                notifyIcon1.Visible = false;//否则该控件不可见
            }
        }

        //双击该控件时触发的事件
        private void notifyIcon1_MouseDoubleClick_1(object sender, MouseEventArgs e)
        {
            Visible = true;//
            WindowState = FormWindowState.Normal;//窗口正常显示
            ShowInTaskbar = true;//在任务栏中显示该窗口
        }

        //private void notifyIcon1_MouseClick_1(object sender, MouseEventArgs e)
        //{
        //    Visible = true;//
        //    WindowState = FormWindowState.Normal;//窗口正常显示
        //    ShowInTaskbar = true;//在任务栏中显示该窗口
        //}

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;   //设置图标不可见
            this.Close();                  //关闭窗体
            this.Dispose();                //释放资源
            Application.Exit();            //关闭应用程序窗体
        }

        private void 显示ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Visible = true;//
            WindowState = FormWindowState.Normal;//窗口正常显示

            ShowInTaskbar = true;//在任务栏中显示该窗口
        }

        //修改关闭窗口的方法，使窗口隐藏
        //  只有Form_Closing事件中 e.Cancel可以用。
        //  你的是Form_Closed事件。 Form_Closed事件时窗口已关了 ，Cancel没用了；
        //  Form_Closing是窗口即将关闭时询问你是不是真的关闭才有Cancel事件
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 注意判断关闭事件reason来源于窗体按钮，否则用菜单退出时无法退出!
            if (e.CloseReason == CloseReason.UserClosing)
            {
                //取消"关闭窗口"事件
                e.Cancel = true; // 取消关闭窗体

                //使关闭时窗口向右下角缩小的效果
                this.WindowState = FormWindowState.Minimized;//窗口最小化
                this.notifyIcon1.Visible = true;
                //this.m_cartoonForm.CartoonClose();
                this.Hide();
                return;
            }

            //e.Cancel = true;
            //WindowState = FormWindowState.Minimized;//窗口最小化
            //this.ShowInTaskbar = false;
            //this.Hide();
        }

        private void lab_ExcelPath_Click(object sender, EventArgs e)
        {
            txt_Excel1.Clear();
        }

        private void lab_WordsPath1_Click(object sender, EventArgs e)
        {
            txt_Words1.Clear();
        }

        private void lab_WordsPath2_Click(object sender, EventArgs e)
        {
            txt_Words2.Clear();
        }

        private void lab_FolderPath_Click(object sender, EventArgs e)
        {
            txt_Folder2.Clear();
        }

        private void modifyFolderConfi_Click(object sender, EventArgs e)
        {
            //string iniFilePath = System.IO.Directory.GetCurrentDirectory() + @"\Data\CreateFolder.ini";
            string iniFilePath = System.Environment.CurrentDirectory + @"\Data\CreateFolder.ini";
            System.Diagnostics.Process.Start("NOTEPAD.exe", iniFilePath);
        }

        private void FrmMain_SizeChanged(object sender, EventArgs e)
        {
            //当窗体最小化时，隐藏到系统托盘。
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.notifyIcon1.Visible = true;
            }
        }

        private void btn_Screencapture_Click(object sender, EventArgs e)
        {
            string screencapture = System.Environment.CurrentDirectory + @"\Screencapture\screencapture.exe";
            Process.Start(screencapture);
        }

        private void 截图OCR识别ToolStripMenuItem_Click(object sender, EventArgs e)
        {          
            string screencapture = System.Environment.CurrentDirectory + @"\Screencapture\screencapture.exe";
            Process.Start(screencapture);
        }

        /// <summary>
        /// 开机自启 复选框的change事件：当单选框选中时，将程序添加至注册表中，未选中时从注册表中移除。并修改配置文件中IsAutoBoot的值。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ck_Selfstart_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                string execPath = Application.ExecutablePath;
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey rk2 = rk.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run");
                if (ck_Selfstart.Checked)
                {
                    rk2.SetValue("GBOfficeTools", execPath);
                }
                else
                {
                    rk2.DeleteValue("GBOfficeTools", false);
                }
                rk2.Close();
                rk.Close();
                LeiTools.ConfigHelper.IniHelper.WriteString("App设置", "IsAutoBoot", Convert.ToString( this.ck_Selfstart.Checked));
                //LeiTools.ConfigHelper.JSONHelper.WriteJSON("IsAutoBoot", ck_Selfstart.Checked);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("[注册表操作]向注册表写开机启动信息失败, Exception: {0}", ex.Message));
            }
        }

        
    }
}