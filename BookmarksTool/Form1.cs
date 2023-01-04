using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace BookmarksTool
{
    public partial class Form1 : Form

    {
        public static Form1 form1; //定义为静态变量，并在构造方法中引用Form1，可以实现在另外的一个类中调用窗体的控件或方法

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
            TextBoxMsg("正在运行……");
            var sw = new Stopwatch();
            sw.Start();  //开始计时
            var start1 = new LeiTools.AsposeOffice.BookmarksReplace();
            start1.StartBookmarksReplace();
            sw.Stop();
            //Console.WriteLine("运行结束,用时{0}秒！按任意键结束", sw.Elapsed);
            Form1.form1.TextBoxMsg("运行结束,用时" + sw.Elapsed + "秒！");
            //MessageBox.Show("Word报告生成完成，请注意查看！");
        }

        private void btn_word2pdf_Click(object sender, EventArgs e)
        {
            TextBoxMsg("正在运行……");
            var sw = new Stopwatch();
            sw.Start(); //开始计时
            //var start2 = new LeiTools.MSOffice.Word.Word2PDF();
            //start2.StartWord2PDF();
            var docPath = Directory.GetCurrentDirectory();
            LeiTools.AsposeOffice.AsposeOfficeWord.Word2PDF(docPath);
            
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
    }
}