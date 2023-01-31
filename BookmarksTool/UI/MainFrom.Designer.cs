namespace BookmarksTool
{
    public partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.uiTabControlMenu1 = new Sunny.UI.UITabControlMenu();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btn_Clear1 = new Sunny.UI.UIImageButton();
            this.btn_start = new Sunny.UI.UISymbolButton();
            this.txt_Words1 = new Sunny.UI.UITextBox();
            this.txt_Excel1 = new Sunny.UI.UITextBox();
            this.uiLabel3 = new Sunny.UI.UILabel();
            this.btn_SelectWords1 = new Sunny.UI.UIButton();
            this.uiLabel2 = new Sunny.UI.UILabel();
            this.but_SelectExcel1 = new Sunny.UI.UIButton();
            this.textBox1 = new Sunny.UI.UITextBox();
            this.uiLabel4 = new Sunny.UI.UILabel();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btn_word2pdf = new Sunny.UI.UISymbolButton();
            this.txt_Folder2 = new Sunny.UI.UITextBox();
            this.txt_Words2 = new Sunny.UI.UITextBox();
            this.uiLabel7 = new Sunny.UI.UILabel();
            this.btn_SelectFolder2 = new Sunny.UI.UIButton();
            this.btn_SelectWords2 = new Sunny.UI.UIButton();
            this.uiLabel8 = new Sunny.UI.UILabel();
            this.textBox2 = new Sunny.UI.UITextBox();
            this.uiLabel6 = new Sunny.UI.UILabel();
            this.btn_Clear2 = new Sunny.UI.UIImageButton();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.btn_ProjectFolder = new Sunny.UI.UISymbolButton();
            this.btn_FileNames = new Sunny.UI.UISymbolButton();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.uiLabel13 = new Sunny.UI.UILabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_Reset = new Sunny.UI.UIButton();
            this.uiLabel11 = new Sunny.UI.UILabel();
            this.txt_bookmarkValueNo = new Sunny.UI.UITextBox();
            this.txt_bookmarkNameNo = new Sunny.UI.UITextBox();
            this.txt_SheetPageNo = new Sunny.UI.UITextBox();
            this.uiLabel10 = new Sunny.UI.UILabel();
            this.uiLabel9 = new Sunny.UI.UILabel();
            this.uiLabel5 = new Sunny.UI.UILabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txt_bookmarkNo = new Sunny.UI.UITextBox();
            this.btn_ChangeSetting = new Sunny.UI.UIButton();
            this.uiLabel12 = new Sunny.UI.UILabel();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lab_Version = new Sunny.UI.UILabel();
            this.btn_OpenHelp = new Sunny.UI.UIButton();
            this.btn_Update = new Sunny.UI.UIButton();
            this.rtxt_help = new Sunny.UI.UIRichTextBox();
            this.uiLabel1 = new Sunny.UI.UILabel();
            this.uiStyleManager1 = new Sunny.UI.UIStyleManager(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.打开ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.uiTabControlMenu1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btn_Clear1)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btn_Clear2)).BeginInit();
            this.tabPage5.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // uiTabControlMenu1
            // 
            this.uiTabControlMenu1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.uiTabControlMenu1.Controls.Add(this.tabPage1);
            this.uiTabControlMenu1.Controls.Add(this.tabPage3);
            this.uiTabControlMenu1.Controls.Add(this.tabPage5);
            this.uiTabControlMenu1.Controls.Add(this.tabPage4);
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
            this.tabPage1.Controls.Add(this.btn_Clear1);
            this.tabPage1.Controls.Add(this.btn_start);
            this.tabPage1.Controls.Add(this.txt_Words1);
            this.tabPage1.Controls.Add(this.txt_Excel1);
            this.tabPage1.Controls.Add(this.uiLabel3);
            this.tabPage1.Controls.Add(this.btn_SelectWords1);
            this.tabPage1.Controls.Add(this.uiLabel2);
            this.tabPage1.Controls.Add(this.but_SelectExcel1);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.uiLabel4);
            this.tabPage1.Location = new System.Drawing.Point(151, 0);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(626, 425);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Word书签工具";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // btn_Clear1
            // 
            this.btn_Clear1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btn_Clear1.BackgroundImage")));
            this.btn_Clear1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btn_Clear1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Clear1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Clear1.Location = new System.Drawing.Point(565, 3);
            this.btn_Clear1.Name = "btn_Clear1";
            this.btn_Clear1.Size = new System.Drawing.Size(45, 29);
            this.btn_Clear1.TabIndex = 26;
            this.btn_Clear1.TabStop = false;
            this.btn_Clear1.Text = null;
            this.btn_Clear1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_Clear1.Click += new System.EventHandler(this.btn_Clear_Click);
            // 
            // btn_start
            // 
            this.btn_start.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_start.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_start.Image = ((System.Drawing.Image)(resources.GetObject("btn_start.Image")));
            this.btn_start.Location = new System.Drawing.Point(237, 372);
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
            // txt_Words1
            // 
            this.txt_Words1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Words1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Words1.Location = new System.Drawing.Point(118, 315);
            this.txt_Words1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Words1.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Words1.Name = "txt_Words1";
            this.txt_Words1.ReadOnly = true;
            this.txt_Words1.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Words1.ShowText = false;
            this.txt_Words1.Size = new System.Drawing.Size(395, 28);
            this.txt_Words1.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Words1.TabIndex = 22;
            this.txt_Words1.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Words1.Watermark = "";
            this.txt_Words1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // txt_Excel1
            // 
            this.txt_Excel1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Excel1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Excel1.Location = new System.Drawing.Point(118, 263);
            this.txt_Excel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Excel1.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Excel1.Name = "txt_Excel1";
            this.txt_Excel1.ReadOnly = true;
            this.txt_Excel1.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Excel1.ShowText = false;
            this.txt_Excel1.Size = new System.Drawing.Size(395, 28);
            this.txt_Excel1.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Excel1.TabIndex = 21;
            this.txt_Excel1.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Excel1.Watermark = "";
            this.txt_Excel1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel3
            // 
            this.uiLabel3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel3.Location = new System.Drawing.Point(8, 266);
            this.uiLabel3.Name = "uiLabel3";
            this.uiLabel3.Size = new System.Drawing.Size(100, 23);
            this.uiLabel3.TabIndex = 18;
            this.uiLabel3.Text = "Excel模板路径：";
            this.uiLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel3.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // btn_SelectWords1
            // 
            this.btn_SelectWords1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_SelectWords1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords1.Location = new System.Drawing.Point(525, 317);
            this.btn_SelectWords1.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_SelectWords1.Name = "btn_SelectWords1";
            this.btn_SelectWords1.Size = new System.Drawing.Size(70, 26);
            this.btn_SelectWords1.TabIndex = 16;
            this.btn_SelectWords1.Text = "选择Words";
            this.btn_SelectWords1.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_SelectWords1.Click += new System.EventHandler(this.btn_SelectWords_Click);
            // 
            // uiLabel2
            // 
            this.uiLabel2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel2.Location = new System.Drawing.Point(261, 6);
            this.uiLabel2.Name = "uiLabel2";
            this.uiLabel2.Size = new System.Drawing.Size(128, 23);
            this.uiLabel2.TabIndex = 15;
            this.uiLabel2.Text = "Word书签工具";
            this.uiLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // but_SelectExcel1
            // 
            this.but_SelectExcel1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.but_SelectExcel1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.but_SelectExcel1.Location = new System.Drawing.Point(525, 265);
            this.but_SelectExcel1.MinimumSize = new System.Drawing.Size(1, 1);
            this.but_SelectExcel1.Name = "but_SelectExcel1";
            this.but_SelectExcel1.Size = new System.Drawing.Size(70, 26);
            this.but_SelectExcel1.TabIndex = 14;
            this.but_SelectExcel1.Text = "选择Excel";
            this.but_SelectExcel1.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.but_SelectExcel1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.but_SelectExcel1.Click += new System.EventHandler(this.but_SelectExcel_Click);
            // 
            // textBox1
            // 
            this.textBox1.AllowDrop = true;
            this.textBox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.textBox1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(22, 40);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox1.MinimumSize = new System.Drawing.Size(1, 16);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.textBox1.ShowScrollBar = true;
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
            this.uiLabel4.Location = new System.Drawing.Point(8, 318);
            this.uiLabel4.Name = "uiLabel4";
            this.uiLabel4.Size = new System.Drawing.Size(111, 23);
            this.uiLabel4.TabIndex = 19;
            this.uiLabel4.Text = "Words文件路径：";
            this.uiLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel4.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // tabPage3
            // 
            this.tabPage3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabPage3.Controls.Add(this.btn_word2pdf);
            this.tabPage3.Controls.Add(this.txt_Folder2);
            this.tabPage3.Controls.Add(this.txt_Words2);
            this.tabPage3.Controls.Add(this.uiLabel7);
            this.tabPage3.Controls.Add(this.btn_SelectFolder2);
            this.tabPage3.Controls.Add(this.btn_SelectWords2);
            this.tabPage3.Controls.Add(this.uiLabel8);
            this.tabPage3.Controls.Add(this.textBox2);
            this.tabPage3.Controls.Add(this.uiLabel6);
            this.tabPage3.Controls.Add(this.btn_Clear2);
            this.tabPage3.Location = new System.Drawing.Point(151, 0);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(626, 425);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Word转PDF";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btn_word2pdf
            // 
            this.btn_word2pdf.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_word2pdf.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.Image = global::BookmarksTool.Properties.Resources.Pdf;
            this.btn_word2pdf.Location = new System.Drawing.Point(237, 372);
            this.btn_word2pdf.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_word2pdf.Name = "btn_word2pdf";
            this.btn_word2pdf.Size = new System.Drawing.Size(158, 29);
            this.btn_word2pdf.Symbol = 61889;
            this.btn_word2pdf.TabIndex = 35;
            this.btn_word2pdf.Text = "Word批量转PDF";
            this.btn_word2pdf.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_word2pdf.Click += new System.EventHandler(this.btn_word2pdf_Click);
            // 
            // txt_Folder2
            // 
            this.txt_Folder2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Folder2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Folder2.Location = new System.Drawing.Point(118, 315);
            this.txt_Folder2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Folder2.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Folder2.Name = "txt_Folder2";
            this.txt_Folder2.ReadOnly = true;
            this.txt_Folder2.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Folder2.ShowText = false;
            this.txt_Folder2.Size = new System.Drawing.Size(395, 28);
            this.txt_Folder2.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Folder2.TabIndex = 34;
            this.txt_Folder2.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Folder2.Watermark = "";
            this.txt_Folder2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // txt_Words2
            // 
            this.txt_Words2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_Words2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_Words2.Location = new System.Drawing.Point(118, 264);
            this.txt_Words2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_Words2.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_Words2.Name = "txt_Words2";
            this.txt_Words2.ReadOnly = true;
            this.txt_Words2.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.txt_Words2.ShowText = false;
            this.txt_Words2.Size = new System.Drawing.Size(395, 28);
            this.txt_Words2.Style = Sunny.UI.UIStyle.Custom;
            this.txt_Words2.TabIndex = 33;
            this.txt_Words2.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.txt_Words2.Watermark = "";
            this.txt_Words2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel7
            // 
            this.uiLabel7.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel7.Location = new System.Drawing.Point(8, 318);
            this.uiLabel7.Name = "uiLabel7";
            this.uiLabel7.Size = new System.Drawing.Size(83, 23);
            this.uiLabel7.TabIndex = 32;
            this.uiLabel7.Text = "文件夹路径：";
            this.uiLabel7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel7.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // btn_SelectFolder2
            // 
            this.btn_SelectFolder2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_SelectFolder2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectFolder2.Location = new System.Drawing.Point(525, 317);
            this.btn_SelectFolder2.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_SelectFolder2.Name = "btn_SelectFolder2";
            this.btn_SelectFolder2.Size = new System.Drawing.Size(70, 26);
            this.btn_SelectFolder2.TabIndex = 30;
            this.btn_SelectFolder2.Text = "选择文件夹";
            this.btn_SelectFolder2.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectFolder2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_SelectFolder2.Click += new System.EventHandler(this.btn_SelectFolder2_Click);
            // 
            // btn_SelectWords2
            // 
            this.btn_SelectWords2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_SelectWords2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords2.Location = new System.Drawing.Point(525, 267);
            this.btn_SelectWords2.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_SelectWords2.Name = "btn_SelectWords2";
            this.btn_SelectWords2.Size = new System.Drawing.Size(70, 26);
            this.btn_SelectWords2.TabIndex = 29;
            this.btn_SelectWords2.Text = "选择Words";
            this.btn_SelectWords2.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_SelectWords2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_SelectWords2.Click += new System.EventHandler(this.btn_SelectWords2_Click);
            // 
            // uiLabel8
            // 
            this.uiLabel8.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel8.Location = new System.Drawing.Point(8, 267);
            this.uiLabel8.Name = "uiLabel8";
            this.uiLabel8.Size = new System.Drawing.Size(111, 23);
            this.uiLabel8.TabIndex = 31;
            this.uiLabel8.Text = "Words文件路径：";
            this.uiLabel8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel8.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // textBox2
            // 
            this.textBox2.AllowDrop = true;
            this.textBox2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.textBox2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox2.Location = new System.Drawing.Point(22, 40);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox2.MinimumSize = new System.Drawing.Size(1, 16);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.RectReadOnlyColor = System.Drawing.Color.FromArgb(((int)(((byte)(80)))), ((int)(((byte)(160)))), ((int)(((byte)(255)))));
            this.textBox2.ShowScrollBar = true;
            this.textBox2.ShowText = false;
            this.textBox2.Size = new System.Drawing.Size(573, 196);
            this.textBox2.Style = Sunny.UI.UIStyle.Custom;
            this.textBox2.TabIndex = 28;
            this.textBox2.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.textBox2.Watermark = "";
            this.textBox2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel6
            // 
            this.uiLabel6.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel6.Location = new System.Drawing.Point(261, 6);
            this.uiLabel6.Name = "uiLabel6";
            this.uiLabel6.Size = new System.Drawing.Size(127, 26);
            this.uiLabel6.TabIndex = 2;
            this.uiLabel6.Text = "Word转PDF";
            this.uiLabel6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel6.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // btn_Clear2
            // 
            this.btn_Clear2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btn_Clear2.BackgroundImage")));
            this.btn_Clear2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btn_Clear2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Clear2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Clear2.Location = new System.Drawing.Point(565, 3);
            this.btn_Clear2.Name = "btn_Clear2";
            this.btn_Clear2.Size = new System.Drawing.Size(45, 29);
            this.btn_Clear2.TabIndex = 27;
            this.btn_Clear2.TabStop = false;
            this.btn_Clear2.Text = null;
            this.btn_Clear2.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_Clear2.Click += new System.EventHandler(this.btn_Clear2_Click);
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.btn_ProjectFolder);
            this.tabPage5.Controls.Add(this.btn_FileNames);
            this.tabPage5.Location = new System.Drawing.Point(151, 0);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(626, 425);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "小工具";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // btn_ProjectFolder
            // 
            this.btn_ProjectFolder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_ProjectFolder.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_ProjectFolder.Location = new System.Drawing.Point(201, 16);
            this.btn_ProjectFolder.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_ProjectFolder.Name = "btn_ProjectFolder";
            this.btn_ProjectFolder.Size = new System.Drawing.Size(146, 35);
            this.btn_ProjectFolder.Style = Sunny.UI.UIStyle.Custom;
            this.btn_ProjectFolder.Symbol = 61564;
            this.btn_ProjectFolder.SymbolColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btn_ProjectFolder.TabIndex = 3;
            this.btn_ProjectFolder.Text = "生成项目文件夹";
            this.btn_ProjectFolder.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_ProjectFolder.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_ProjectFolder.Click += new System.EventHandler(this.btn_ProjectFolder_Click);
            // 
            // btn_FileNames
            // 
            this.btn_FileNames.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_FileNames.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_FileNames.Location = new System.Drawing.Point(28, 16);
            this.btn_FileNames.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_FileNames.Name = "btn_FileNames";
            this.btn_FileNames.Size = new System.Drawing.Size(146, 35);
            this.btn_FileNames.Style = Sunny.UI.UIStyle.Custom;
            this.btn_FileNames.Symbol = 261787;
            this.btn_FileNames.SymbolColor = System.Drawing.Color.Aqua;
            this.btn_FileNames.TabIndex = 2;
            this.btn_FileNames.Text = "获取文件名";
            this.btn_FileNames.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_FileNames.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_FileNames.Click += new System.EventHandler(this.btn_FileNames_Click);
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.uiLabel13);
            this.tabPage4.Controls.Add(this.groupBox2);
            this.tabPage4.Controls.Add(this.txt_bookmarkValueNo);
            this.tabPage4.Controls.Add(this.txt_bookmarkNameNo);
            this.tabPage4.Controls.Add(this.txt_SheetPageNo);
            this.tabPage4.Controls.Add(this.uiLabel10);
            this.tabPage4.Controls.Add(this.uiLabel9);
            this.tabPage4.Controls.Add(this.uiLabel5);
            this.tabPage4.Controls.Add(this.groupBox1);
            this.tabPage4.Location = new System.Drawing.Point(151, 0);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(626, 425);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "软件设置";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // uiLabel13
            // 
            this.uiLabel13.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel13.Location = new System.Drawing.Point(258, 6);
            this.uiLabel13.Name = "uiLabel13";
            this.uiLabel13.Size = new System.Drawing.Size(127, 26);
            this.uiLabel13.TabIndex = 8;
            this.uiLabel13.Text = "软件设置";
            this.uiLabel13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel13.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_Reset);
            this.groupBox2.Controls.Add(this.uiLabel11);
            this.groupBox2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox2.Location = new System.Drawing.Point(319, 42);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(272, 251);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Excel模板书签设置说明";
            // 
            // btn_Reset
            // 
            this.btn_Reset.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Reset.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Reset.Location = new System.Drawing.Point(90, 202);
            this.btn_Reset.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_Reset.Name = "btn_Reset";
            this.btn_Reset.Size = new System.Drawing.Size(100, 35);
            this.btn_Reset.TabIndex = 6;
            this.btn_Reset.Text = "恢复默认";
            this.btn_Reset.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Reset.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_Reset.Click += new System.EventHandler(this.btn_Reset_Click);
            // 
            // uiLabel11
            // 
            this.uiLabel11.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.uiLabel11.Font = new System.Drawing.Font("阿里巴巴普惠体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel11.Location = new System.Drawing.Point(6, 26);
            this.uiLabel11.Name = "uiLabel11";
            this.uiLabel11.Size = new System.Drawing.Size(260, 166);
            this.uiLabel11.TabIndex = 0;
            this.uiLabel11.Text = "1. 书签所在Sheet页：sheet1 = 0，sheet2 = 1，依次类推；默认为Sheet1（0）。\r\n2. 书签名\\值所在列：A列 = 0，B列 = 1" +
    "，依次类推；默认书签名为D列（3）、书签值为C列（2）。\r\n3. 书签个数：默认为300，可根据项目需要调整。";
            this.uiLabel11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel11.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // txt_bookmarkValueNo
            // 
            this.txt_bookmarkValueNo.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_bookmarkValueNo.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_bookmarkValueNo.Location = new System.Drawing.Point(194, 150);
            this.txt_bookmarkValueNo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_bookmarkValueNo.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_bookmarkValueNo.Name = "txt_bookmarkValueNo";
            this.txt_bookmarkValueNo.ShowText = false;
            this.txt_bookmarkValueNo.Size = new System.Drawing.Size(57, 29);
            this.txt_bookmarkValueNo.TabIndex = 4;
            this.txt_bookmarkValueNo.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.txt_bookmarkValueNo.Watermark = "";
            this.txt_bookmarkValueNo.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.txt_bookmarkValueNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_bookmarkValueNo_KeyPress);
            // 
            // txt_bookmarkNameNo
            // 
            this.txt_bookmarkNameNo.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_bookmarkNameNo.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_bookmarkNameNo.Location = new System.Drawing.Point(194, 111);
            this.txt_bookmarkNameNo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_bookmarkNameNo.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_bookmarkNameNo.Name = "txt_bookmarkNameNo";
            this.txt_bookmarkNameNo.ShowText = false;
            this.txt_bookmarkNameNo.Size = new System.Drawing.Size(57, 29);
            this.txt_bookmarkNameNo.TabIndex = 3;
            this.txt_bookmarkNameNo.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.txt_bookmarkNameNo.Watermark = "";
            this.txt_bookmarkNameNo.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.txt_bookmarkNameNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_bookmarkNameNo_KeyPress);
            // 
            // txt_SheetPageNo
            // 
            this.txt_SheetPageNo.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_SheetPageNo.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_SheetPageNo.Location = new System.Drawing.Point(194, 72);
            this.txt_SheetPageNo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_SheetPageNo.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_SheetPageNo.Name = "txt_SheetPageNo";
            this.txt_SheetPageNo.ShowText = false;
            this.txt_SheetPageNo.Size = new System.Drawing.Size(57, 29);
            this.txt_SheetPageNo.TabIndex = 1;
            this.txt_SheetPageNo.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.txt_SheetPageNo.Watermark = "";
            this.txt_SheetPageNo.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.txt_SheetPageNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txt_SheetPageNo_KeyPress);
            // 
            // uiLabel10
            // 
            this.uiLabel10.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel10.Location = new System.Drawing.Point(42, 153);
            this.uiLabel10.Name = "uiLabel10";
            this.uiLabel10.Size = new System.Drawing.Size(140, 23);
            this.uiLabel10.TabIndex = 0;
            this.uiLabel10.Text = "书签值所在列：";
            this.uiLabel10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel10.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel9
            // 
            this.uiLabel9.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel9.Location = new System.Drawing.Point(42, 114);
            this.uiLabel9.Name = "uiLabel9";
            this.uiLabel9.Size = new System.Drawing.Size(140, 23);
            this.uiLabel9.TabIndex = 0;
            this.uiLabel9.Text = "书签名所在列：";
            this.uiLabel9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel9.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // uiLabel5
            // 
            this.uiLabel5.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel5.Location = new System.Drawing.Point(42, 75);
            this.uiLabel5.Name = "uiLabel5";
            this.uiLabel5.Size = new System.Drawing.Size(140, 23);
            this.uiLabel5.TabIndex = 0;
            this.uiLabel5.Text = "书签所在Sheet页：";
            this.uiLabel5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel5.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt_bookmarkNo);
            this.groupBox1.Controls.Add(this.btn_ChangeSetting);
            this.groupBox1.Controls.Add(this.uiLabel12);
            this.groupBox1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox1.Location = new System.Drawing.Point(23, 42);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(272, 251);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel模板书签设置";
            // 
            // txt_bookmarkNo
            // 
            this.txt_bookmarkNo.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.txt_bookmarkNo.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txt_bookmarkNo.Location = new System.Drawing.Point(171, 147);
            this.txt_bookmarkNo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txt_bookmarkNo.MinimumSize = new System.Drawing.Size(1, 16);
            this.txt_bookmarkNo.Name = "txt_bookmarkNo";
            this.txt_bookmarkNo.ShowText = false;
            this.txt_bookmarkNo.Size = new System.Drawing.Size(57, 29);
            this.txt_bookmarkNo.TabIndex = 9;
            this.txt_bookmarkNo.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.txt_bookmarkNo.Watermark = "";
            this.txt_bookmarkNo.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // btn_ChangeSetting
            // 
            this.btn_ChangeSetting.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_ChangeSetting.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_ChangeSetting.Location = new System.Drawing.Point(89, 202);
            this.btn_ChangeSetting.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_ChangeSetting.Name = "btn_ChangeSetting";
            this.btn_ChangeSetting.Size = new System.Drawing.Size(100, 35);
            this.btn_ChangeSetting.TabIndex = 5;
            this.btn_ChangeSetting.Text = "修改设置";
            this.btn_ChangeSetting.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_ChangeSetting.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_ChangeSetting.Click += new System.EventHandler(this.btn_ChangeSetting_Click);
            // 
            // uiLabel12
            // 
            this.uiLabel12.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.uiLabel12.Location = new System.Drawing.Point(19, 150);
            this.uiLabel12.Name = "uiLabel12";
            this.uiLabel12.Size = new System.Drawing.Size(140, 23);
            this.uiLabel12.TabIndex = 8;
            this.uiLabel12.Text = "书签个数：";
            this.uiLabel12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel12.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lab_Version);
            this.tabPage2.Controls.Add(this.btn_OpenHelp);
            this.tabPage2.Controls.Add(this.btn_Update);
            this.tabPage2.Controls.Add(this.rtxt_help);
            this.tabPage2.Controls.Add(this.uiLabel1);
            this.tabPage2.Location = new System.Drawing.Point(151, 0);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(626, 425);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "关于软件";
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
            // btn_OpenHelp
            // 
            this.btn_OpenHelp.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_OpenHelp.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_OpenHelp.Location = new System.Drawing.Point(169, 360);
            this.btn_OpenHelp.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_OpenHelp.Name = "btn_OpenHelp";
            this.btn_OpenHelp.Size = new System.Drawing.Size(92, 30);
            this.btn_OpenHelp.TabIndex = 3;
            this.btn_OpenHelp.Text = "软件说明";
            this.btn_OpenHelp.TipsFont = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_OpenHelp.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.btn_OpenHelp.Click += new System.EventHandler(this.btn_OpenHelp_Click);
            // 
            // btn_Update
            // 
            this.btn_Update.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_Update.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Update.Location = new System.Drawing.Point(319, 360);
            this.btn_Update.MinimumSize = new System.Drawing.Size(1, 1);
            this.btn_Update.Name = "btn_Update";
            this.btn_Update.Size = new System.Drawing.Size(92, 30);
            this.btn_Update.TabIndex = 3;
            this.btn_Update.Text = "检查更新";
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
            this.uiLabel1.Text = "关于软件";
            this.uiLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.uiLabel1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "GBOfficeTools";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick_1);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.打开ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(101, 48);
            this.contextMenuStrip1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.contextMain);
            // 
            // 打开ToolStripMenuItem
            // 
            this.打开ToolStripMenuItem.Name = "打开ToolStripMenuItem";
            this.打开ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.打开ToolStripMenuItem.Text = "显示";
            this.打开ToolStripMenuItem.Click += new System.EventHandler(this.显示ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // MainForm
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
            this.Name = "MainForm";
            this.ShowRadius = false;
            this.ShowRect = false;
            this.Text = "GBOfficeTools";
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 650, 450);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.uiTabControlMenu1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btn_Clear1)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btn_Clear2)).EndInit();
            this.tabPage5.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
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
        private Sunny.UI.UIButton but_SelectExcel1;
        private Sunny.UI.UILabel uiLabel2;
        private Sunny.UI.UIButton btn_SelectWords1;
        private Sunny.UI.UITextBox txt_Excel1;
        private Sunny.UI.UILabel uiLabel3;
        private Sunny.UI.UILabel uiLabel4;
        private Sunny.UI.UITextBox txt_Words1;
        private Sunny.UI.UISymbolButton btn_start;
        private Sunny.UI.UIImageButton btn_Clear1;
        private Sunny.UI.UIButton btn_Update;
        private Sunny.UI.UILabel lab_Version;
        private System.Windows.Forms.TabPage tabPage3;
        private Sunny.UI.UILabel uiLabel6;
        private Sunny.UI.UISymbolButton btn_word2pdf;
        private Sunny.UI.UITextBox txt_Folder2;
        private Sunny.UI.UITextBox txt_Words2;
        private Sunny.UI.UILabel uiLabel7;
        private Sunny.UI.UIButton btn_SelectFolder2;
        private Sunny.UI.UIButton btn_SelectWords2;
        private Sunny.UI.UILabel uiLabel8;
        private Sunny.UI.UITextBox textBox2;
        private Sunny.UI.UIImageButton btn_Clear2;
        private Sunny.UI.UIButton btn_OpenHelp;
        private System.Windows.Forms.TabPage tabPage4;
        private Sunny.UI.UILabel uiLabel5;
        private Sunny.UI.UILabel uiLabel10;
        private Sunny.UI.UILabel uiLabel9;
        private Sunny.UI.UITextBox txt_bookmarkValueNo;
        private Sunny.UI.UITextBox txt_bookmarkNameNo;
        private Sunny.UI.UITextBox txt_SheetPageNo;
        private System.Windows.Forms.GroupBox groupBox2;
        private Sunny.UI.UILabel uiLabel11;
        private System.Windows.Forms.GroupBox groupBox1;
        private Sunny.UI.UIButton btn_ChangeSetting;
        private Sunny.UI.UIButton btn_Reset;
        private Sunny.UI.UITextBox txt_bookmarkNo;
        private Sunny.UI.UILabel uiLabel12;
        private Sunny.UI.UILabel uiLabel13;
        private System.Windows.Forms.TabPage tabPage5;
        private Sunny.UI.UISymbolButton btn_FileNames;
        private Sunny.UI.UISymbolButton btn_ProjectFolder;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 打开ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
    }
}

