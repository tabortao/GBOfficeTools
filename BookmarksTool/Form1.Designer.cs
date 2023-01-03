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
            this.label1 = new System.Windows.Forms.Label();
            this.btn_word2pdf = new System.Windows.Forms.Button();
            this.btn_help = new System.Windows.Forms.Button();
            this.btn_start = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.White;
            this.textBox1.Location = new System.Drawing.Point(29, 101);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(527, 218);
            this.textBox1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(232, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 22);
            this.label1.TabIndex = 2;
            this.label1.Text = "绿建海绵工具箱";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // btn_word2pdf
            // 
            this.btn_word2pdf.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_word2pdf.Image = global::BookmarksTool.Properties.Resources.Pdf;
            this.btn_word2pdf.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_word2pdf.Location = new System.Drawing.Point(321, 348);
            this.btn_word2pdf.Name = "btn_word2pdf";
            this.btn_word2pdf.Size = new System.Drawing.Size(145, 32);
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
            this.btn_help.Location = new System.Drawing.Point(491, 1);
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
            this.btn_start.Location = new System.Drawing.Point(99, 348);
            this.btn_start.Name = "btn_start";
            this.btn_start.Size = new System.Drawing.Size(155, 32);
            this.btn_start.TabIndex = 0;
            this.btn_start.Text = "批量生成Word报告";
            this.btn_start.UseVisualStyleBackColor = true;
            this.btn_start.Click += new System.EventHandler(this.Btn_start_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(584, 411);
            this.Controls.Add(this.btn_word2pdf);
            this.Controls.Add(this.btn_help);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_start);
            this.Controls.Add(this.textBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "BookmarksTool V2.0";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btn_start;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_help;
        private System.Windows.Forms.Button btn_word2pdf;
    }
}

