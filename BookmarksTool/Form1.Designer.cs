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
            this.btn_word2pdf = new System.Windows.Forms.Button();
            this.btn_start = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_Excel = new System.Windows.Forms.TextBox();
            this.txt_Words = new System.Windows.Forms.TextBox();
            this.but_SelectExcel = new System.Windows.Forms.Button();
            this.btn_SelectWords = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_Folder = new System.Windows.Forms.TextBox();
            this.btn_SelectFolder = new System.Windows.Forms.Button();
            this.uiTabControlMenu1 = new Sunny.UI.UITabControlMenu();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_help = new System.Windows.Forms.Button();
            this.uiLabel1 = new Sunny.UI.UILabel();
            this.textBox1 = new Sunny.UI.UITextBox();
            this.rtxt_help = new Sunny.UI.UIRichTextBox();
            this.uiTabControlMenu1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_word2pdf
            // 
            this.btn_word2pdf.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.Image = global::BookmarksTool.Properties.Resources.Pdf;
            this.btn_word2pdf.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_word2pdf.Location = new System.Drawing.Point(353, 364);
            this.btn_word2pdf.Name = "btn_word2pdf";
            this.btn_word2pdf.Size = new System.Drawing.Size(153, 32);
            this.btn_word2pdf.TabIndex = 4;
            this.btn_word2pdf.Text = "Word批量转PDF";
            this.btn_word2pdf.UseVisualStyleBackColor = true;
            this.btn_word2pdf.Click += new System.EventHandler(this.btn_word2pdf_Click);
            // 
            // btn_start
            // 
            this.btn_start.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_start.Image = global::BookmarksTool.Properties.Resources.start;
            this.btn_start.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_start.Location = new System.Drawing.Point(129, 364);
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
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(15, 254);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Excel模板路径：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(15, 287);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 17);
            this.label3.TabIndex = 5;
            this.label3.Text = "Words文件路径：";
            // 
            // txt_Excel
            // 
            this.txt_Excel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Excel.Location = new System.Drawing.Point(116, 252);
            this.txt_Excel.Name = "txt_Excel";
            this.txt_Excel.ReadOnly = true;
            this.txt_Excel.Size = new System.Drawing.Size(424, 21);
            this.txt_Excel.TabIndex = 6;
            // 
            // txt_Words
            // 
            this.txt_Words.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Words.Location = new System.Drawing.Point(116, 285);
            this.txt_Words.Name = "txt_Words";
            this.txt_Words.ReadOnly = true;
            this.txt_Words.Size = new System.Drawing.Size(424, 21);
            this.txt_Words.TabIndex = 7;
            // 
            // but_SelectExcel
            // 
            this.but_SelectExcel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.but_SelectExcel.Location = new System.Drawing.Point(558, 251);
            this.but_SelectExcel.Name = "but_SelectExcel";
            this.but_SelectExcel.Size = new System.Drawing.Size(74, 23);
            this.but_SelectExcel.TabIndex = 8;
            this.but_SelectExcel.Text = "选择Excel";
            this.but_SelectExcel.UseVisualStyleBackColor = true;
            this.but_SelectExcel.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_SelectWords
            // 
            this.btn_SelectWords.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords.Location = new System.Drawing.Point(558, 284);
            this.btn_SelectWords.Name = "btn_SelectWords";
            this.btn_SelectWords.Size = new System.Drawing.Size(74, 23);
            this.btn_SelectWords.TabIndex = 9;
            this.btn_SelectWords.Text = "选择Words";
            this.btn_SelectWords.UseVisualStyleBackColor = true;
            this.btn_SelectWords.Click += new System.EventHandler(this.btn_SelectWords_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(15, 317);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 17);
            this.label4.TabIndex = 10;
            this.label4.Text = "文件夹路径：";
            // 
            // txt_Folder
            // 
            this.txt_Folder.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Folder.Location = new System.Drawing.Point(116, 315);
            this.txt_Folder.Name = "txt_Folder";
            this.txt_Folder.ReadOnly = true;
            this.txt_Folder.Size = new System.Drawing.Size(424, 21);
            this.txt_Folder.TabIndex = 11;
            // 
            // btn_SelectFolder
            // 
            this.btn_SelectFolder.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectFolder.Location = new System.Drawing.Point(558, 314);
            this.btn_SelectFolder.Name = "btn_SelectFolder";
            this.btn_SelectFolder.Size = new System.Drawing.Size(74, 23);
            this.btn_SelectFolder.TabIndex = 12;
            this.btn_SelectFolder.Text = "选择文件夹";
            this.btn_SelectFolder.UseVisualStyleBackColor = true;
            this.btn_SelectFolder.Click += new System.EventHandler(this.btn_SelectFolder_Click);
            // 
            // uiTabControlMenu1
            // 
            this.uiTabControlMenu1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.uiTabControlMenu1.Controls.Add(this.tabPage1);
            this.uiTabControlMenu1.Controls.Add(this.tabPage2);
            this.uiTabControlMenu1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.uiTabControlMenu1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTabControlMenu1.Location = new System.Drawing.Point(-4, 35);
            this.uiTabControlMenu1.Multiline = true;
            this.uiTabControlMenu1.Name = "uiTabControlMenu1";
            this.uiTabControlMenu1.SelectedIndex = 0;
            this.uiTabControlMenu1.Size = new System.Drawing.Size(861, 424);
            this.uiTabControlMenu1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.uiTabControlMenu1.Style = Sunny.UI.UIStyle.Custom;
            this.uiTabControlMenu1.TabIndex = 13;
            this.uiTabControlMenu1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btn_SelectFolder);
            this.tabPage1.Controls.Add(this.txt_Folder);
            this.tabPage1.Controls.Add(this.btn_start);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.btn_help);
            this.tabPage1.Controls.Add(this.btn_SelectWords);
            this.tabPage1.Controls.Add(this.btn_word2pdf);
            this.tabPage1.Controls.Add(this.but_SelectExcel);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.txt_Words);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.txt_Excel);
            this.tabPage1.Location = new System.Drawing.Point(201, 0);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(660, 424);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "绿建海绵工具箱";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.rtxt_help);
            this.tabPage2.Controls.Add(this.uiLabel1);
            this.tabPage2.Location = new System.Drawing.Point(201, 0);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(660, 424);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "软件使用帮助";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.Click += new System.EventHandler(this.tabPage2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(259, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 22);
            this.label1.TabIndex = 2;
            this.label1.Text = "绿建海绵工具箱";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // btn_help
            // 
            this.btn_help.FlatAppearance.BorderSize = 0;
            this.btn_help.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_help.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_help.Image = global::BookmarksTool.Properties.Resources.help;
            this.btn_help.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_help.Location = new System.Drawing.Point(548, 2);
            this.btn_help.Name = "btn_help";
            this.btn_help.Size = new System.Drawing.Size(102, 29);
            this.btn_help.TabIndex = 3;
            this.btn_help.Text = "使用说明";
            this.btn_help.UseVisualStyleBackColor = true;
            this.btn_help.Click += new System.EventHandler(this.Btn_help_Click);
            // 
            // uiLabel1
            // 
            this.uiLabel1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel1.Location = new System.Drawing.Point(258, 6);
            this.uiLabel1.Name = "uiLabel1";
            this.uiLabel1.Size = new System.Drawing.Size(127, 26);
            this.uiLabel1.TabIndex = 1;
            this.uiLabel1.Text = "软件使用帮助";
            this.uiLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // textBox1
            // 
            this.textBox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.textBox1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(18, 40);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox1.MinimumSize = new System.Drawing.Size(1, 16);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ShowText = false;
            this.textBox1.Size = new System.Drawing.Size(614, 196);
            this.textBox1.TabIndex = 13;
            this.textBox1.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.textBox1.Watermark = "";
            this.textBox1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // rtxt_help
            // 
            this.rtxt_help.FillColor = System.Drawing.Color.White;
            this.rtxt_help.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtxt_help.Location = new System.Drawing.Point(21, 37);
            this.rtxt_help.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rtxt_help.MinimumSize = new System.Drawing.Size(1, 1);
            this.rtxt_help.Name = "rtxt_help";
            this.rtxt_help.Padding = new System.Windows.Forms.Padding(2);
            this.rtxt_help.ShowText = false;
            this.rtxt_help.Size = new System.Drawing.Size(610, 347);
            this.rtxt_help.Style = Sunny.UI.UIStyle.Custom;
            this.rtxt_help.TabIndex = 2;
            this.rtxt_help.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.rtxt_help.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(850, 450);
            this.Controls.Add(this.uiTabControlMenu1);
            this.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(850, 450);
            this.MinimumSize = new System.Drawing.Size(820, 450);
            this.Name = "Form1";
            this.Style = Sunny.UI.UIStyle.Custom;
            this.Text = "BookmarksTool V2.4";
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 650, 450);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.uiTabControlMenu1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btn_start;
        private System.Windows.Forms.Button btn_word2pdf;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_Excel;
        private System.Windows.Forms.TextBox txt_Words;
        private System.Windows.Forms.Button but_SelectExcel;
        private System.Windows.Forms.Button btn_SelectWords;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_Folder;
        private System.Windows.Forms.Button btn_SelectFolder;
        private Sunny.UI.UITabControlMenu uiTabControlMenu1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_help;
        private Sunny.UI.UILabel uiLabel1;
        private Sunny.UI.UITextBox textBox1;
        private Sunny.UI.UIRichTextBox rtxt_help;
    }
}

