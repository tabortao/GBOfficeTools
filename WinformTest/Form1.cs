using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WinformTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_folderBrowser_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择需要转换的Word文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var savePath = dialog.SelectedPath;
                richTextBox1.Text = savePath;
                //textBox1.Text = savePath;
            }
        }
    }
}
