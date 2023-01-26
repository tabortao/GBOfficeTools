using System;
using System.Windows.Forms;
using WinformTest.ConfigHelper;

namespace WinformTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = IniHelper.ReadString("软件设置", "书签所在Sheet页", "NA");
            textBox2.Text = IniHelper.ReadString("软件设置", "书签名所在列", "NA");
            textBox3.Text = IniHelper.ReadString("软件设置", "书签值所在列", "NA");
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //IniHelper.WriteString("节","sheet号码","1");
            //IniHelper.WriteString("节", "sheet号码2", "2");

            IniHelper.WriteString("软件设置", "书签所在Sheet页", textBox1.Text);
            IniHelper.WriteString("软件设置", "书签名所在列", textBox2.Text);
            IniHelper.WriteString("软件设置", "书签值所在列", textBox3.Text);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
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
    }
}