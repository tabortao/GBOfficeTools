using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace BookmarksTool
{
    public partial class Form1 : Form

    {
        public static Form1 form1; //定义为静态变量，并在构造方法中引用Form1，可以实现在另外的一个类中调用窗体的控件或方法
        public static string excelFilePath; // 定义静态变量 excelFilePath，用于存储excel模板文件路径
        public static string[] wordsPath;// 定义静态变量 wordsPath，用于存储选中的多个word模板文件路径

        public Form1()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            form1 = this; //构造方法中引用Form1
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
            //Console.WriteLine(@"正在运行……");
            //textBox1.Text += "正在运行……";
            //string mCode = RegInfo.GetMachineCode();
            //string leiCode = "681348742402123950127693";
            textBox1.Text = ""; //清空文本框内容，为下次执行扫除干净
            if (String.IsNullOrEmpty(txt_Excel.Text))
                TextBoxMsg("正在运行……");
            var sw = new Stopwatch();
            sw.Start();  //开始计时
            var excelName = Path.GetFileName(excelFilePath);
            TextBoxMsg("当前计算采用模板为：" + excelName);
            if (txt_Excel.Text.Length != 0 && txt_Words.Text.Length != 0)
            {
                LeiTools.AsposeOffice.BookmarksReplace.ReportMaker(excelFilePath, wordsPath);
            }
            else if (txt_Excel.Text.Length == 0 && txt_Words.Text.Length == 0)
            {
                //如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
                LeiTools.AsposeOffice.BookmarksReplace.ReportMaker();
            }
            else if (txt_Excel.Text.Length == 0 || txt_Words.Text.Length == 0)
            {
                MessageBox.Show("请选择Excel模板文件或选择Word文件");
            }

            //LeiTools.AsposeOffice.BookmarksReplace.StartBookmarksReplace2();
            sw.Stop();
            //Console.WriteLine("运行结束,用时{0}秒！按任意键结束", sw.Elapsed);
            Form1.form1.TextBoxMsg("运行结束,用时" + sw.Elapsed + "秒！");
            //MessageBox.Show("Word报告生成完成，请注意查看！");
        }

        private void btn_word2pdf_Click(object sender, EventArgs e)
        {
            textBox1.Text = ""; //清空文本框内容，为下次执行扫除干净
            TextBoxMsg("正在运行……");
            var sw = new Stopwatch();
            sw.Start(); //开始计时
            //var start2 = new LeiTools.MSOffice.Word.Word2PDF();
            //start2.StartWord2PDF();          
            if (txt_Words.Text.Length != 0)
            {
                LeiTools.AsposeOffice.AsposeOfficeWord.Word2PDF(wordsPath);
            }
            else if (txt_Words.Text.Length == 0)
            {
                //如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
                LeiTools.AsposeOffice.AsposeOfficeWord.Word2PDF();
            }
            else
            {
                MessageBox.Show("选择的Word文件有误，请重试！");
            }
            

            sw.Stop();
            Form1.form1.TextBoxMsg("运行结束，用时" + sw.Elapsed + "秒");
            //MessageBox.Show("Word批量转换PDF完成，请注意查看！");
        }

        private void Btn_help_Click(object sender, EventArgs e)
        {
            var help =
                "    1. 在excel内填写书签名和书签内容，书签英文名；\r\n" +
                "    2. 在Word中添加书签；\r\n" +
                "    3. 将含有书签内容的excel和要批量替换的Word文件放入同一个文件夹内；\r\n" +
                "    4. 执行BookmarksTool.exe，即可迅速替换完成。\r\n" +
                "    5. 作者：筑博姚蕾。\r\n" +
                "    6. 如有问题，欢迎联系，邮箱：yaoleistable@qq.com。\r\n";
            MessageBox.Show(help, "BookmarksTool 使用说明");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenExcelFileDialog();
        }

        private void btn_SelectWords_Click(object sender, EventArgs e)
        {
            OpenWordFileDialog();
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
                    TextBoxMsg(wordPath);
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
                TextBoxMsg(excelFilePath);
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
    }
}