
namespace WinformTest
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btn_folderBrowser = new System.Windows.Forms.Button();
            this.btn_openFileDialog = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_folderBrowser
            // 
            this.btn_folderBrowser.Location = new System.Drawing.Point(66, 129);
            this.btn_folderBrowser.Name = "btn_folderBrowser";
            this.btn_folderBrowser.Size = new System.Drawing.Size(128, 23);
            this.btn_folderBrowser.TabIndex = 0;
            this.btn_folderBrowser.Text = "打开文件夹对话框";
            this.btn_folderBrowser.UseVisualStyleBackColor = true;
            this.btn_folderBrowser.Click += new System.EventHandler(this.btn_folderBrowser_Click);
            // 
            // btn_openFileDialog
            // 
            this.btn_openFileDialog.Location = new System.Drawing.Point(238, 129);
            this.btn_openFileDialog.Name = "btn_openFileDialog";
            this.btn_openFileDialog.Size = new System.Drawing.Size(128, 23);
            this.btn_openFileDialog.TabIndex = 3;
            this.btn_openFileDialog.Text = "打开选择文件对话框";
            this.btn_openFileDialog.UseVisualStyleBackColor = true;
            this.btn_openFileDialog.Click += new System.EventHandler(this.btn_openFileDialog_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(66, 25);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(658, 80);
            this.textBox1.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btn_openFileDialog);
            this.Controls.Add(this.btn_folderBrowser);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btn_folderBrowser;
        private System.Windows.Forms.Button btn_openFileDialog;
        private System.Windows.Forms.TextBox textBox1;
    }
}

