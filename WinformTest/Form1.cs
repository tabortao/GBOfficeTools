using System;
using System.Windows.Forms;

namespace WinformTest
{
    public partial class Form1 : Form
    {
        public static Form1 form1;

        public Form1()
        {
            InitializeComponent();
        }
        private string FolderDialog(string name)
        {
            string FolderPath = name;
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择需要转换的Word文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //获取选择的目录路径
                FolderPath = dialog.SelectedPath;
            }
            return FolderPath;
        }


        private void btn_folderBrowser_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择需要转换的Word文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var savePath = dialog.SelectedPath;
                TextBoxMsg(savePath);
                //textBox1.Text = savePath;
            }
        }

        public void TextBoxMsg(string msg)  //输出消息到多行文本框
        {
            textBox1.AppendText(msg + "\r\n");
            //textBox1.Text += msg + "\r\n";
        }

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
                string[] filesPath = dialog.FileNames;
                foreach ( var filePath in filesPath)
                {
                    TextBoxMsg(filePath);                  
                }

                //获取选择的文件路径
                //string str = "";
                //for (int i = 0; i < dialog.FileNames.Length; i++) //根据数组长度定义循环次数
                //    str += dialog.FileNames.GetValue(i).ToString();//获取文件文件名
                //TextBoxMsg(str);
                //MessageBox.Show(String.Format("Name={0} !", str));
            }
            //return FilePath;
        }

        private void btn_openFileDialog_Click(object sender, EventArgs e)
        {
            OpenWordFileDialog();
            //OpenFileDialog dlg = new OpenFileDialog(); //创建一个OpenFileDialog
            //dlg.Multiselect = true;          //设置属性为多选
            //if (dlg.ShowDialog() == DialogResult.OK)
            //{
            //    string str = "";
            //    for (int i = 0; i < dlg.FileNames.Length; i++) //根据数组长度定义循环次数
            //        str += dlg.FileNames.GetValue(i).ToString();//获取文件文件名
            //    TextBoxMsg(str);
            //    MessageBox.Show(String.Format("Name={0} !", str));
            //}
        }
    }
}