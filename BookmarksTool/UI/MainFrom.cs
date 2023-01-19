using Sunny.UI;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace BookmarksTool
{
    public partial class Form1 : UIForm

    {
        public static Form1 form1; //定义为静态变量，并在构造方法中引用Form1，可以实现在另外的一个类中调用窗体的控件或方法
        public static string excelFilePath; // 定义静态变量 excelFilePath，用于存储excel模板文件路径
        public static string[] wordsPath;// 定义静态变量 wordsPath，用于存储选中的多个word模板文件路径
        public static string folderPath; // 定义静态变量 folderPath，用于存储文件夹路径

        public string help =
                "    一、批量生成Word使用说明：\r\n" +
                "       （1）方式一：选择Excel文件，选择要转换的Words文件，点击“批量生成Word”按钮。\r\n" +
                "       （2）方式二：拷贝软件到存放Excel和Word的文件夹，点击“批量生成Word”按钮。\r\n" +
                "    二、Word批量转PDF使用说明：\r\n" +
                "       （1）方式一：选择要转换的Words文件，点击“Word批量转PDF”按钮。\r\n" +
                "       （2）方式二：选择文件夹路径（存放Word的文件夹）,点击“Word批量转PDF”按钮\r\n" +
                "       （3）方式三：拷贝软件到存放Excel和Word的文件夹，点击“Word批量转PDF”按钮。\r\n" +
                "    三、关于软件：\r\n" +
                "    （1）作者：筑博设计@绿色建筑部@姚蕾。\r\n" +
                "    （2）如有问题，欢迎联系，微信：yao-lei\r\n";

        public Form1()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            form1 = this; //构造方法中引用Form1
            this.rtxt_help.Text = help;
            this.lab_Version.Text = "当前版本：" + Application.ProductVersion.ToString() + "\n";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //打开软件后，自动检查是否有可用更新,如有更新，弹出对话框询问是否更新，点击更新后，自动更新。
            LeiTools.AutoUpdate.AutoCheckUpdate();
        }

        public void TextBoxMsg(string msg)  //输出消息到多行文本框
        {
            textBox1.AppendText(msg + "\r\n");
            //textBox1.Text += msg + "\r\n";
        }

        private void label1_Click(object sender, EventArgs e)
        {
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
                if (txt_Excel.Text.Length != 0 && txt_Words.Text.Length != 0)
                {
                    //普通循环运算，进行word书签替换
                    LeiTools.AsposeOffice.BookmarksReplace.ReportMaker(excelFilePath, wordsPath);
                    //并行运算，进行word书签替换
                    //LeiTools.AsposeOffice.BookmarksReplace.ParallelReportMaker(excelFilePath, wordsPath);
                }
                else if (txt_Excel.Text.Length == 0 && txt_Words.Text.Length == 0)
                {
                    //如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
                    //普通循环运算，进行word书签替换
                    LeiTools.AsposeOffice.BookmarksReplace.ReportMaker();
                    //并行运算，进行word书签替换
                    //LeiTools.AsposeOffice.BookmarksReplace.ParallelReportMaker();
                }
                else if (txt_Excel.Text.Length == 0 || txt_Words.Text.Length == 0)
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
                Form1.form1.TextBoxMsg("运行结束,用时" + sw.Elapsed + "秒！");

                #endregion 执行Word报告生成
            }
            else
            {
                MessageBox.Show("软件授权过期，请关注微信公众号“可持续学园”留言，免费获取更新。");
            }
        }

        private void btn_word2pdf_Click(object sender, EventArgs e)
        {
            //textBox1.ResetText();
            textBox1.Clear();
            //textBox1.Text = "";
            //System.Threading.Thread.Sleep(1000);

            //textBox1.Text = ""; //清空文本框内容，为下次执行扫除干净
            TextBoxMsg("Word批量转PDF正在运行，请稍等……");
            var sw = new Stopwatch();
            sw.Start(); //开始计时
            //var start2 = new LeiTools.MSOffice.Word.Word2PDF();
            //start2.StartWord2PDF();
            if (txt_Words.Text.Length != 0 || (txt_Folder.Text.Length != 0 && txt_Words.Text.Length != 0))
            {
                //如果选择了多个word文件，或者既选择了文件夹，又选择了多个word文件时，将多个word文件转为PDF
                LeiTools.AsposeOffice.AsposeOfficeWord.Word2PDF(wordsPath);
            }
            else if (txt_Folder.Text.Length != 0)
            {
                //如果选择了文件夹路径，把文件夹内的word转为pdf
                LeiTools.AsposeOffice.AsposeOfficeWord.Word2PDF(folderPath);
            }
            else if (txt_Words.Text.Length == 0 && txt_Folder.Text.Length == 0)
            {
                //如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
                LeiTools.AsposeOffice.AsposeOfficeWord.Word2PDF();
            }
            else
            {
                MessageBox.Show("选择的Word文件有误，请重试！");
            }
            sw.Stop();//计时结束
            Form1.form1.TextBoxMsg("Word批量转PDF已完成，用时" + sw.Elapsed + "秒");
        }

        /// <summary>
        /// 打开word文件选择对话框
        /// </summary>
        private void OpenWordFileDialog()
        {
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
                    txt_Words.AppendText(wordPath + "\r\n");
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
        /// 打开excel文件选中对话框
        /// </summary>
        private void OpenExcelFileDialog()
        {
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
                txt_Excel.AppendText(excelFilePath);

                //获取选择的文件路径
                //string str = "";
                //for (int i = 0; i < dialog.FileNames.Length; i++) //根据数组长度定义循环次数
                //    str += dialog.FileNames.GetValue(i).ToString();//获取文件文件名
                //TextBoxMsg(str);
                //MessageBox.Show(String.Format("Name={0} !", str));
            }
            //return FilePath;
        }

        private void btn_SelectFolder_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件夹";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                folderPath = dialog.SelectedPath;
                //TextBoxMsg(folderPath);
                txt_Folder.AppendText(folderPath);
            }
        }

        private void uiRichTextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void but_SelectExcel_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            OpenExcelFileDialog();
        }

        private void btn_SelectWords_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            OpenWordFileDialog();
        }

        /// <summary>
        /// 清除所有文本框内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Clear_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            txt_Excel.Clear();
            txt_Words.Clear();
            txt_Folder.Clear();
        }

        private void btn_Update_Click(object sender, EventArgs e)
        {
            //调用自动更新函数，进行软件更新
            LeiTools.AutoUpdate.CheckUpdate();
        }
    }
}