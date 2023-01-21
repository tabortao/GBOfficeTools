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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.uiTabControlMenu1 = new Sunny.UI.UITabControlMenu();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btn_Clear = new Sunny.UI.UIImageButton();
            this.btn_word2pdf = new Sunny.UI.UISymbolButton();
            this.btn_start = new Sunny.UI.UISymbolButton();
            this.txt_Folder = new Sunny.UI.UITextBox();
            this.txt_Words = new Sunny.UI.UITextBox();
            this.txt_Excel = new Sunny.UI.UITextBox();
            this.uiLabel5 = new Sunny.UI.UILabel();
            this.uiLabel3 = new Sunny.UI.UILabel();
            this.btn_SelectFolder = new Sunny.UI.UIButton();
            this.btn_SelectWords = new Sunny.UI.UIButton();
            this.uiLabel2 = new Sunny.UI.UILabel();
            this.but_SelectExcel = new Sunny.UI.UIButton();
            this.textBox1 = new Sunny.UI.UITextBox();
            this.uiLabel4 = new Sunny.UI.UILabel();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lab_Version = new Sunny.UI.UILabel();
            this.btn_Update = new Sunny.UI.UIButton();
            this.rtxt_help = new Sunny.UI.UIRichTextBox();
            this.uiLabel1 = new Sunny.UI.UILabel();
            this.uiStyleManager1 = new Sunny.UI.UIStyleManager(this.components);
            this.uiTabControlMenu1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btn_Clear)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // uiTabControlMenu1
            // 
            this.uiTabControlMenu1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.uiTabControlMenu1.Controls.Add(this.tabPage1);
            this.uiTabControlMenu1.Controls.Add(this.tabPage2);
            this.uiTabControlMenu1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.uiTabControlMenu1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiTabControlMenu1.ItemSize = new System.Drawing.Size(150, 40);
            this.uiTabControlMenu1.Location = new System.Drawing.Point(-4, 35);
            this.uiTabControlMenu1.Multiline = true;
            this.uiTabControlMenu1.Name = "uiTabControlMenu1";
            this.uiTabControlMenu1.SelectedIndex = 0;
            this.uiTabControlMenu1.Size = new System.Drawing.Size(777, 425);
            this.uiTabControlMenu1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.uiTabControlMenu1.TabIndex = 13;
            this.uiTabControlMenu1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btn_Clear);
            this.tabPage1.Controls.Add(this.btn_word2pdf);
            this.tabPage1.Controls.Add(this.btn_start);
            this.tabPage1.Controls.Add(this.txt_Folder);
            this.tabPage1.Controls.Add(this.txt_Words);
            this.tabPage1.Controls.Add(this.txt_Excel);
            this.tabPage1.Controls.Add(this.uiLabel5);
            this.tabPage1.Controls.Add(this.uiLabel3);
            this.tabPage1.Controls.Add(this.btn_SelectFolder);
            this.tabPage1.Controls.Add(this.btn_SelectWords);
            this.tabPage1.Controls.Add(this.uiLabel2);
            this.tabPage1.Controls.Add(this.but_SelectExcel);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.uiLabel4);
            this.tabPage1.Location = new System.Drawing.Point(151, 0);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(626, 425);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "绿建海绵工具箱";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btn_Clear
            // 
            this.btn_Clear.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btn_Clear.BackgroundImage")));
            this.btn_Clear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btn_Clear.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Clear.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Clear.Location = new System.Drawing.Point(565, 3);
            this.btn_Clear.Name = "btn_Clear";
            this.btn_Clear.Size = new System.Drawing.Size(45, 29);
            this.btn_Clear.TabIndex = 26;
            this.btn_Clear.TabStop = false;
            this.btn_Clear.Text = null;
            this.btn_Clear.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_Clear.Click += new System.EventHandler(this.btn_Clear_Click);
            // 
            // btn_word2pdf
            // 
            this.btn_word2pdf.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_word2pdf.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.Image = global::BookmarksTool.Properties.Resources.Pdf;
            this.btn_word2pdf.Location = new System.Drawing.Point(344, 372);
            this.btn_word2pdf.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_word2pdf.Name = "btn_word2pdf";
            this.btn_word2pdf.Size = new System.Drawing.Size(157, 29);
            this.btn_word2pdf.Symbol = 61889;
            this.btn_word2pdf.TabIndex = 25;
            this.btn_word2pdf.Text = "Word批量转PDF";
            this.btn_word2pdf.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_word2pdf.Click += new System.EventHandler(this.btn_word2pdf_Click);
            // 
            // btn_start
            // 
            this.btn_start.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_start.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_start.Image = ((System.Drawing.Image)(resources.GetObject("btn_start.Image")));
            this.btn_start.Location = new System.Drawing.Point(127, 372);
            this.btn_start.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_start.Name = "btn_start";
            this.btn_start.Size = new System.Drawing.Size(158, 29);
            this.btn_start.Symbol = 73;
            this.btn_start.TabIndex = 24;
            this.btn_start.Text = "批量生成Word报告";
            this.btn_start.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_start.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_start.Click += new System.EventHandler(this.Btn_start_Click);
            // 
            // txt_Folder
            // 
            this.txt_Folder.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Folder.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Folder.Location = new System.Drawing.Point(118, 327);
            this.txt_Folder.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Folder.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Folder.Name = "txt_Folder";
            this.txt_Folder.ReadOnly = true;
            this.txt_Folder.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Folder.ShowText = false;
            this.txt_Folder.Size = new System.Drawing.Size(395, 28);
            this.txt_Folder.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Folder.TabIndex = 23;
            this.txt_Folder.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Folder.Watermark = "";
            this.txt_Folder.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // txt_Words
            // 
            this.txt_Words.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Words.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Words.Location = new System.Drawing.Point(118, 287);
            this.txt_Words.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Words.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Words.Name = "txt_Words";
            this.txt_Words.ReadOnly = true;
            this.txt_Words.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Words.ShowText = false;
            this.txt_Words.Size = new System.Drawing.Size(395, 28);
            this.txt_Words.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Words.TabIndex = 22;
            this.txt_Words.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Words.Watermark = "";
            this.txt_Words.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // txt_Excel
            // 
            this.txt_Excel.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Excel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Excel.Location = new System.Drawing.Point(118, 247);
            this.txt_Excel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Excel.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Excel.Name = "txt_Excel";
            this.txt_Excel.ReadOnly = true;
            this.txt_Excel.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Excel.ShowText = false;
            this.txt_Excel.Size = new System.Drawing.Size(395, 28);
            this.txt_Excel.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Excel.TabIndex = 21;
            this.txt_Excel.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Excel.Watermark = "";
            this.txt_Excel.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel5
            // 
            this.uiLabel5.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel5.Location = new System.Drawing.Point(8, 330);
            this.uiLabel5.Name = "uiLabel5";
            this.uiLabel5.Size = new System.Drawing.Size(83, 23);
            this.uiLabel5.TabIndex = 20;
            this.uiLabel5.Text = "文件夹路径：";
            this.uiLabel5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel5.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel3
            // 
            this.uiLabel3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel3.Location = new System.Drawing.Point(8, 250);
            this.uiLabel3.Name = "uiLabel3";
            this.uiLabel3.Size = new System.Drawing.Size(100, 23);
            this.uiLabel3.TabIndex = 18;
            this.uiLabel3.Text = "Excel模板路径：";
            this.uiLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel3.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // btn_SelectFolder
            // 
            this.btn_SelectFolder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_SelectFolder.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectFolder.Location = new System.Drawing.Point(525, 329);
            this.btn_SelectFolder.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_SelectFolder.Name = "btn_SelectFolder";
            this.btn_SelectFolder.Size = new System.Drawing.Size(70, 26);
            this.btn_SelectFolder.TabIndex = 17;
            this.btn_SelectFolder.Text = "选择文件夹";
            this.btn_SelectFolder.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectFolder.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_SelectFolder.Click += new System.EventHandler(this.btn_SelectFolder_Click);
            // 
            // btn_SelectWords
            // 
            this.btn_SelectWords.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_SelectWords.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords.Location = new System.Drawing.Point(525, 289);
            this.btn_SelectWords.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_SelectWords.Name = "btn_SelectWords";
            this.btn_SelectWords.Size = new System.Drawing.Size(70, 26);
            this.btn_SelectWords.TabIndex = 16;
            this.btn_SelectWords.Text = "选择Words";
            this.btn_SelectWords.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_SelectWords.Click += new System.EventHandler(this.btn_SelectWords_Click);
            // 
            // uiLabel2
            // 
            this.uiLabel2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel2.Location = new System.Drawing.Point(261, 6);
            this.uiLabel2.Name = "uiLabel2";
            this.uiLabel2.Size = new System.Drawing.Size(128, 23);
            this.uiLabel2.TabIndex = 15;
            this.uiLabel2.Text = "绿建海绵工具箱";
            this.uiLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // but_SelectExcel
            // 
            this.but_SelectExcel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.but_SelectExcel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.but_SelectExcel.Location = new System.Drawing.Point(525, 249);
            this.but_SelectExcel.MinimumSize = new System.Drawing.Size(1, 1);
            this.but_SelectExcel.Name = "but_SelectExcel";
            this.but_SelectExcel.Size = new System.Drawing.Size(70, 26);
            this.but_SelectExcel.TabIndex = 14;
            this.but_SelectExcel.Text = "选择Excel";
            this.but_SelectExcel.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.but_SelectExcel.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.but_SelectExcel.Click += new System.EventHandler(this.but_SelectExcel_Click);
            // 
            // textBox1
            // 
            this.textBox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.textBox1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(22, 40);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox1.MinimumSize = new System.Drawing.Size(1, 16);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.textBox1.ShowText = false;
            this.textBox1.Size = new System.Drawing.Size(573, 196);
            this.textBox1.Style = Sunny.UI.UIStyle.Custom;
            this.textBox1.TabIndex = 13;
            this.textBox1.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.textBox1.Watermark = "";
            this.textBox1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel4
            // 
            this.uiLabel4.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel4.Location = new System.Drawing.Point(8, 290);
            this.uiLabel4.Name = "uiLabel4";
            this.uiLabel4.Size = new System.Drawing.Size(111, 23);
            this.uiLabel4.TabIndex = 19;
            this.uiLabel4.Text = "Words文件路径：";
            this.uiLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel4.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lab_Version);
            this.tabPage2.Controls.Add(this.btn_Update);
            this.tabPage2.Controls.Add(this.rtxt_help);
            this.tabPage2.Controls.Add(this.uiLabel1);
            this.tabPage2.Location = new System.Drawing.Point(151, 0);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(626, 425);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "软件使用帮助";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lab_Version
            // 
            this.lab_Version.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lab_Version.Location = new System.Drawing.Point(466, 388);
            this.lab_Version.Name = "lab_Version";
            this.lab_Version.Size = new System.Drawing.Size(141, 23);
            this.lab_Version.TabIndex = 4;
            this.lab_Version.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lab_Version.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // btn_Update
            // 
            this.btn_Update.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Update.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Update.Location = new System.Drawing.Point(262, 354);
            this.btn_Update.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(88, 30);
            this.btn_Update.TabIndex = 3;
            this.btn_Update.Text = "立即更新";
            this.btn_Update.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Update.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_Update.Click += new System.EventHandler(this.btn_Update_Click);
            // 
            // rtxt_help
            // 
            this.rtxt_help.FillColor = System.Drawing.Color.White;
            this.rtxt_help.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtxt_help.Location = new System.Drawing.Point(19, 41);
            this.rtxt_help.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rtxt_help.MinimumSize = new System.Drawing.Size(1, 1);
            this.rtxt_help.Name = "rtxt_help";
            this.rtxt_help.Padding = new System.Windows.Forms.Padding(2);
            this.rtxt_help.ReadOnly = true;
            this.rtxt_help.ShowText = false;
            this.rtxt_help.Size = new System.Drawing.Size(577, 292);
            this.rtxt_help.Style = Sunny.UI.UIStyle.Custom;
            this.rtxt_help.TabIndex = 2;
            this.rtxt_help.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.rtxt_help.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
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
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(249)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(760, 450);
            this.Controls.Add(this.uiTabControlMenu1);
            this.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(760, 450);
            this.MinimumSize = new System.Drawing.Size(760, 450);
            this.Name = "Form1";
            this.ShowRadius = false;
            this.ShowRect = false;
            this.Text = "BookmarksTool 2";
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 650, 450);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.uiTabControlMenu1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btn_Clear)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private Sunny.UI.UITabControlMenu uiTabControlMenu1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private Sunny.UI.UILabel uiLabel1;
        private Sunny.UI.UITextBox textBox1;
        private Sunny.UI.UIRichTextBox rtxt_help;
        private Sunny.UI.UIStyleManager uiStyleManager1;
        private Sunny.UI.UIButton but_SelectExcel;
        private Sunny.UI.UILabel uiLabel2;
        private Sunny.UI.UIButton btn_SelectWords;
        private Sunny.UI.UIButton btn_SelectFolder;
        private Sunny.UI.UITextBox txt_Excel;
        private Sunny.UI.UILabel uiLabel5;
        private Sunny.UI.UILabel uiLabel3;
        private Sunny.UI.UILabel uiLabel4;
        private Sunny.UI.UITextBox txt_Words;
        private Sunny.UI.UITextBox txt_Folder;
        private Sunny.UI.UISymbolButton btn_word2pdf;
        private Sunny.UI.UISymbolButton btn_start;
        private Sunny.UI.UIImageButton btn_Clear;
        private Sunny.UI.UIButton btn_Update;
        private Sunny.UI.UILabel lab_Version;
    }
}

