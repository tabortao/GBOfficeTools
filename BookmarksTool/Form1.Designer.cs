namespace BookmarksTool
{
    public partial class Form1
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btn_word2pdf = new System.Windows.Forms.Button();
            this.btn_help = new System.Windows.Forms.Button();
            this.btn_start = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_Excel = new System.Windows.Forms.TextBox();
            this.txt_Words = new System.Windows.Forms.TextBox();
            this.but_SelectExcel = new System.Windows.Forms.Button();
            this.btn_SelectWords = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_Folder = new System.Windows.Forms.TextBox();
            this.btn_SelectFolder = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.White;
            this.textBox1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(28, 54);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(581, 203);
            this.textBox1.TabIndex = 1;
            // 
            // btn_word2pdf
            // 
            this.btn_word2pdf.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.Image = global::BookmarksTool.Properties.Resources.Pdf;
            this.btn_word2pdf.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_word2pdf.Location = new System.Drawing.Point(352, 372);
            this.btn_word2pdf.Name = "btn_word2pdf";
            this.btn_word2pdf.Size = new System.Drawing.Size(153, 32);
            this.btn_word2pdf.TabIndex = 4;
            this.btn_word2pdf.Text = "Word批量转PDF";
            this.btn_word2pdf.UseVisualStyleBackColor = true;
            this.btn_word2pdf.Click += new System.EventHandler(this.btn_word2pdf_Click);
            // 
            // btn_help
            // 
            this.btn_help.FlatAppearance.BorderSize = 0;
            this.btn_help.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_help.Image = global::BookmarksTool.Properties.Resources.help;
            this.btn_help.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_help.Location = new System.Drawing.Point(544, 1);
            this.btn_help.Name = "btn_help";
            this.btn_help.Size = new System.Drawing.Size(89, 29);
            this.btn_help.TabIndex = 3;
            this.btn_help.Text = "使用说明";
            this.btn_help.UseVisualStyleBackColor = true;
            this.btn_help.Click += new System.EventHandler(this.Btn_help_Click);
            // 
            // btn_start
            // 
            this.btn_start.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_start.Image = global::BookmarksTool.Properties.Resources.start;
            this.btn_start.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_start.Location = new System.Drawing.Point(128, 372);
            this.btn_start.Name = "btn_start";
            this.btn_start.Size = new System.Drawing.Size(153, 32);
            this.btn_start.TabIndex = 0;
            this.btn_start.Text = "批量生成Word报告";
            this.btn_start.UseVisualStyleBackColor = true;
            this.btn_start.Click += new System.EventHandler(this.Btn_start_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 278);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "Excel模板路径：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 310);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(95, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "Words文件路径：";
            // 
            // txt_Excel
            // 
            this.txt_Excel.Location = new System.Drawing.Point(114, 274);
            this.txt_Excel.Name = "txt_Excel";
            this.txt_Excel.ReadOnly = true;
            this.txt_Excel.Size = new System.Drawing.Size(410, 21);
            this.txt_Excel.TabIndex = 6;
            // 
            // txt_Words
            // 
            this.txt_Words.Location = new System.Drawing.Point(114, 306);
            this.txt_Words.Name = "txt_Words";
            this.txt_Words.ReadOnly = true;
            this.txt_Words.Size = new System.Drawing.Size(410, 21);
            this.txt_Words.TabIndex = 7;
            // 
            // but_SelectExcel
            // 
            this.but_SelectExcel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.but_SelectExcel.Location = new System.Drawing.Point(534, 273);
            this.but_SelectExcel.Name = "but_SelectExcel";
            this.but_SelectExcel.Size = new System.Drawing.Size(75, 23);
            this.but_SelectExcel.TabIndex = 8;
            this.but_SelectExcel.Text = "选择Excel";
            this.but_SelectExcel.UseVisualStyleBackColor = true;
            this.but_SelectExcel.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_SelectWords
            // 
            this.btn_SelectWords.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords.Location = new System.Drawing.Point(534, 302);
            this.btn_SelectWords.Name = "btn_SelectWords";
            this.btn_SelectWords.Size = new System.Drawing.Size(75, 23);
            this.btn_SelectWords.TabIndex = 9;
            this.btn_SelectWords.Text = "选择Words";
            this.btn_SelectWords.UseVisualStyleBackColor = true;
            this.btn_SelectWords.Click += new System.EventHandler(this.btn_SelectWords_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(261, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 22);
            this.label1.TabIndex = 2;
            this.label1.Text = "绿建海绵工具箱";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 344);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "文件夹路径：";
            // 
            // txt_Folder
            // 
            this.txt_Folder.Location = new System.Drawing.Point(114, 340);
            this.txt_Folder.Name = "txt_Folder";
            this.txt_Folder.ReadOnly = true;
            this.txt_Folder.Size = new System.Drawing.Size(410, 21);
            this.txt_Folder.TabIndex = 11;
            // 
            // btn_SelectFolder
            // 
            this.btn_SelectFolder.Location = new System.Drawing.Point(534, 339);
            this.btn_SelectFolder.Name = "btn_SelectFolder";
            this.btn_SelectFolder.Size = new System.Drawing.Size(75, 23);
            this.btn_SelectFolder.TabIndex = 12;
            this.btn_SelectFolder.Text = "选择文件夹";
            this.btn_SelectFolder.UseVisualStyleBackColor = true;
            this.btn_SelectFolder.Click += new System.EventHandler(this.btn_SelectFolder_Click);
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(634, 411);
            this.Controls.Add(this.btn_SelectFolder);
            this.Controls.Add(this.txt_Folder);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btn_SelectWords);
            this.Controls.Add(this.but_SelectExcel);
            this.Controls.Add(this.txt_Words);
            this.Controls.Add(this.txt_Excel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_word2pdf);
            this.Controls.Add(this.btn_help);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_start);
            this.Controls.Add(this.textBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(650, 450);
            this.MinimumSize = new System.Drawing.Size(650, 450);
            this.Name = "Form1";
            this.Text = "BookmarksTool V2.3";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btn_start;
        private System.Windows.Forms.Button btn_help;
        private System.Windows.Forms.Button btn_word2pdf;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_Excel;
        private System.Windows.Forms.TextBox txt_Words;
        private System.Windows.Forms.Button but_SelectExcel;
        private System.Windows.Forms.Button btn_SelectWords;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_Folder;
        private System.Windows.Forms.Button btn_SelectFolder;
    }
}

