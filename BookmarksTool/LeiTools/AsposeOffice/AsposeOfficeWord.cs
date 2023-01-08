using Aspose.Words;
using System;
using System.Drawing;
using System.IO;

namespace BookmarksTool.LeiTools.AsposeOffice

{
    public class AsposeOfficeWord
    {
        //在word模板中设置书签，相同内容使用交叉引用，更新域后会自动改
        //本class使用了Aspose.Words.dll工具

        private Aspose.Words.DocumentBuilder builder;
        private Aspose.Words.Document doc; //文档对象

        /// <summary>
        /// 载入模版
        /// </summary>
        /// <param name="strFileName">文件名路径</param>
        public void Open(string strFileName)
        {
            if (!string.IsNullOrEmpty(strFileName))
            {
                doc = new Aspose.Words.Document(strFileName);
            }
        }

        public void Open()
        {
            doc = new Aspose.Words.Document();
        }

        public void Builder()
        {
            builder = new Aspose.Words.DocumentBuilder(doc);
        }

        /// <summary>
        /// 替换书签内容
        /// </summary>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="value">替换内容</param>
        public void ReplaceBookMark(string bookmarkName, string value, string type = "", double width = 320, double height = 320)
        {
            var bm = doc.Range.Bookmarks[bookmarkName];
            if (bm == null)
            {
                return;
            }
            try
            {
                if (type == "IMG")
                {
                    bm.Text = "";
                    builder.MoveToBookmark(bookmarkName);
                    var img = builder.InsertImage(@value, width, height); // width 350像素， height 350像素
                                                                          //Console.WriteLine(bookmarkName+"替换图片成功！");
                    doc.UpdateFields(); //更新域
                }
                else
                {
                    bm.Text = "";
                    builder.MoveToBookmark(bookmarkName);
                    //bm.Text = value;//会把文字格式清空
                    builder.Write(value); //修改书签内容的第二种方法，不会更改文字字体等属性
                                          //Console.WriteLine(bookmarkName+"替换书签内容成功！");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("错误：" + e.Message);
            }
        }

        public void ReplaceBookMark(string bookmarkName, Bitmap value, string type = "", double width = 320, double height = 320)
        {
            var bm = doc.Range.Bookmarks[bookmarkName];
            if (bm == null)
            {
                return;
            }
            try
            {
                if (type == "IMG")
                {
                    bm.Text = "";
                    builder.MoveToBookmark(bookmarkName);
                    var img = builder.InsertImage(@value, width, height); // width 350像素， height 350像素
                                                                          //Console.WriteLine(bookmarkName+"替换图片成功！");
                    doc.UpdateFields(); //更新域
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("错误：" + e.Message);
            }
        }

        /// <summary>
        /// 删除所有书签
        /// </summary>
        public void DeleteAllBookmark()
        {
            doc.Range.Bookmarks.Clear();
        }

        /// <summary>
        /// 打印书签
        /// </summary>
        public void PrintBookmarkName()
        {
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine("书签名：" + bookmark.Name);
                Console.WriteLine("书签内容：" + bookmark.Text);
            }
        }

        /// <summary>
        /// 设置字体
        /// </summary>
        /// <param name="fontName"></param>
        public void SetFont(string fontName)
        {
            Aspose.Words.Font font = builder.Font;
            font.NameFarEast = fontName; //设置字体
                                         //font.Size = 22; //设置字体大小
        }

        /// <summary>
        /// 写入√ or ×
        /// </summary>
        /// <param name="box">box=0,×；box=1，√</param>
        /// <returns></returns>
        public void WriteBox(int box = 0)
        {
            SetFont("Wingdings 2");
            if (box == 0)
            { builder.Write("\u00A3"); }
            else if (box == 1)
            {
                builder.Write("\u00A3");
            }
        }

        public void Save(string strFileName)
        {
            doc.UpdateFields(); //更新域
            doc.Save(strFileName);
        }

        #region 批量word转pdf

        /// <summary>
        /// Word2PDF()无参数时，直接将软件所在文件夹内的word转换为PDF。
        /// </summary>
        /// <returns></returns>
        public static bool Word2PDF()
        {
            var wordPath = Directory.GetCurrentDirectory();
            bool result = false;
            //实现查找路径中word文件,带来筛选，直接选出word文件。
            var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //SaveToPDF(wordFiles);
            ParallelSaveToPDF(wordFiles); // 并行循环执行Word转PDF
            //application.Quit();//退出word
            return result;
        }

        /// <summary>
        /// 把某个文件夹内的全部word转换为PDF
        /// </summary>
        /// <param name="wordPath">word文件夹路径</param>
        /// <returns></returns>
        public static bool Word2PDF(string wordPath)
        {
            bool result = false;
            //实现查找路径中word文件,带来筛选，直接选出word文件。
            var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //SaveToPDF(wordFiles);
            ParallelSaveToPDF(wordFiles); // 并行循环执行Word转PDF
            //application.Quit();//退出word
            return result;
        }

        /// <summary>
        /// 选中多个word文件，然后转换为PDF
        /// </summary>
        /// <param name="wordFiles">word文件路径，数组格式</param>
        /// <returns></returns>
        public static bool Word2PDF(string[] wordFiles)
        {
            bool result = false;
            //实现查找路径中word文件,带来筛选，直接选出word文件。
            //var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //var wordFiles = Directory.EnumerateFiles(wordPath, "*.doc");
            //SaveToPDF(wordFiles);
            ParallelSaveToPDF(wordFiles); // 并行循环执行Word转PDF
            //application.Quit();//退出word
            return result;
        }

        /// <summary>
        /// 普通循环，批量将Word转PDF
        /// </summary>
        /// <param name="wordFiles"></param>
        /// <returns></returns>
        public static bool SaveToPDF(string[] wordFiles)
        {
            bool result = false;
            foreach (var wordFile in wordFiles)
            {
                string wordFileFolder = Path.GetDirectoryName(wordFile);
                string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFile); //获取文件名称，不含拓展名。
                string pdfFilePath = wordFileFolder + @"\" + wordFileNameWithoutExtension + ".pdf"; //设置pdf文件存储路径。
                //循环，转换每一个word文件。
                Document doc = new Document(wordFile);
                try
                {
                    if (!File.Exists(pdfFilePath))
                    {
                        SaveToPDF(wordFile, pdfFilePath);
                        result = true;
                    }
                    else
                    {
                        File.Delete(pdfFilePath);
                        SaveToPDF(wordFile, pdfFilePath);
                        result = true;
                    }
                }
                catch (Exception e)
                {
                    //System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                    //Console.WriteLine(e.Message);
                    Form1.form1.TextBoxMsg(e.Message);
                    result = false;
                }
            }
            return result;
        }

        /// <summary>
        /// 并行循环，批量将Word转PDF
        /// </summary>
        /// <param name="wordFiles"></param>
        /// <returns></returns>
        public static bool ParallelSaveToPDF(string[] wordFiles)
        {
            bool result = false;

            //循环，转换每一个word文件。
            System.Threading.Tasks.Parallel.ForEach(wordFiles, wordFile =>
            {
                string wordFileFolder = Path.GetDirectoryName(wordFile);
                string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFile); //获取文件名称，不含拓展名。
                string pdfFilePath = wordFileFolder + @"\" + wordFileNameWithoutExtension + ".pdf"; //设置pdf文件存储路径。
                Document doc = new Document(wordFile);
                try
                {
                    if (!File.Exists(pdfFilePath))
                    {
                        SaveToPDF(wordFile, pdfFilePath);
                        result = true;
                    }
                    else
                    {
                        File.Delete(pdfFilePath);
                        SaveToPDF(wordFile, pdfFilePath);
                        result = true;
                    }
                }
                catch (Exception e)
                {
                    //System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                    //Console.WriteLine(e.Message);
                    Form1.form1.TextBoxMsg(e.Message);
                    result = false;
                }
            });

            return result;
        }

        /// <summary>
        /// 单个word文件转换为pdf
        /// </summary>
        /// <param name="wordFilePath">word文件路径，含文件名及拓展名</param>
        /// <param name="pdfFilePath">pdf文件路径，含文件名及拓展名</param>
        public static bool SaveToPDF(string wordFilePath, string pdfFilePath)
        {
            bool result;
            try
            {
                string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFilePath); //获取文件名称，不含拓展名。
                Aspose.Words.Document doc = new Aspose.Words.Document(wordFilePath);
                doc.Save(pdfFilePath, Aspose.Words.SaveFormat.Pdf);
                Form1.form1.TextBoxMsg(wordFileNameWithoutExtension + "转PDF成功!");
                return true;
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        #endregion 批量word转pdf
    }
}