using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

// 本程序使用引用-添加引用-扩展 Microsoft.Office.Interop.Word.dll，要求系统需安装office
// 本程序已测试成功，为控制台应用程序，将程序放入包含word的文件夹内，执行本程序，即可实现把文件夹内所有word转换为pdf
//WordTOPDF.exe
//1.系统需安装微软office软件才可以运行；
//2.测试转换38个word文件用时1分51.05s；（采用并行计算，来优化速度）
//3.转换质量较高，暂无问题。

namespace BookmarksTool.LeiTools.MSOffice.Word
{
    /// <summary>
    /// 通过调用微软office，实现批量转换
    /// </summary>
    public class Word2PDF
    {
        //internal object w;

        /// <summary>
        /// Word转Pdf核心函数
        /// </summary>
        private bool WordTOPDF(DirectoryInfo docPath)
        {
            bool result = false;
            Application application = new Application
            {
                Visible = false
            };
            Document document = null;
            FileInfo[] files = docPath.GetFiles();

            #region 并行循环

            Parallel.ForEach(files, f =>
            {
                string sourcePath = f.FullName;

                if (f.Extension == ".docx")
                {
                    try
                    {
                        document = application.Documents.Open(sourcePath);
                        //document.Activate();
                        string PDFPath = sourcePath.Replace(".docx", ".pdf");//pdf存放位置
                        if (!File.Exists(@PDFPath))//存在PDF，不需要继续转换
                        {
                            document.ExportAsFixedFormat(PDFPath, WdExportFormat.wdExportFormatPDF);
                            Form1.form1.TextBoxMsg("文件{0}转换PDF成功" + f.Name);
                            //Console.WriteLine("文件{0}转换PDF成功", f.Name);
                        }
                        result = true;
                    }
                    catch (Exception e)
                    {
                        //Form1.form1.TextBoxMsg("请关闭需要转换的所有word文档。"+ "\r\n" + e.Message);
                        //Console.WriteLine(e.Message);
                        System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);

                        result = false;
                    }
                    finally
                    {
                        if (document != null)
                        {
                            document.Close();
                            //Marshal.ReleaseComObject(document);
                            //document = null;
                        }
                        if (application != null)
                        {
                            application.Quit();
                        }
                    }
                }
                else if (f.Extension == ".doc")
                {
                    try
                    {
                        document = application.Documents.Open(sourcePath);
                        string PDFPath = sourcePath.Replace(".doc", ".pdf");//pdf存放位置
                        if (!File.Exists(@PDFPath))//存在PDF，不需要继续转换
                        {
                            document.ExportAsFixedFormat(PDFPath, WdExportFormat.wdExportFormatPDF);
                            Form1.form1.TextBoxMsg("文件{0}转换PDF成功" + f.Name);
                            //Console.WriteLine("文件{0}转换PDF成功", f.Name);
                        }
                        result = true;
                    }
                    catch (Exception e)
                    {
                        //Console.WriteLine(e.Message);
                        //Form1.form1.TextBoxMsg("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                        System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                        result = false;
                    }
                    finally
                    {
                        if (document != null)
                        {
                            document.Close();
                            Marshal.ReleaseComObject(document);
                            //document = null;
                        }
                        if (application != null)
                        {
                            application.Quit();
                            Marshal.ReleaseComObject(application);
                            //application = null;
                        }
                    }
                }
            });

            #endregion 并行循环

            #region 普通循环

            /*
            foreach (FileInfo f in files)
            {
                string sourcePath = f.FullName;
                if (f.Extension == ".docx")
                {
                    try
                    {
                        document = application.Documents.Open(sourcePath);
                        string PDFPath = sourcePath.Replace(".docx", ".pdf");//pdf存放位置
                        if (!File.Exists(@PDFPath))//存在PDF，不需要继续转换
                        {
                            document.ExportAsFixedFormat(PDFPath, WdExportFormat.wdExportFormatPDF);
                            Console.WriteLine("文件{0}转换PDF成功", f.Name);
                        }
                        result = true;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        result = false;
                    }
                    finally
                    {
                        document.Close();
                    }
                }
                else if (f.Extension == ".doc")
                {
                    try
                    {
                        document = application.Documents.Open(sourcePath);
                        string PDFPath = sourcePath.Replace(".doc", ".pdf");//pdf存放位置
                        if (!File.Exists(@PDFPath))//存在PDF，不需要继续转换
                        {
                            document.ExportAsFixedFormat(PDFPath, WdExportFormat.wdExportFormatPDF);
                            Console.WriteLine("文件{0}转换PDF成功", f.Name);
                        }
                        result = true;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        result = false;
                    }
                    finally
                    {
                        document.Close();
                    }
                }
            }
            */

            #endregion 普通循环

            return result;
        }

        public void StartWord2PDF()
        {
            var docPath = Directory.GetCurrentDirectory();
            var path = new DirectoryInfo(docPath);
            var w = new BookmarksTool.LeiTools.MSOffice.Word.Word2PDF();
            w.WordTOPDF(path);
        }

        //class Program
        //{
        //    static void Main(string[] args)
        //    {
        //        Stopwatch sw = new Stopwatch();
        //        sw.Start();
        //        string docPath = Directory.GetCurrentDirectory();

        //        Console.WriteLine("当前路径为：{0}", docPath);
        //        //string docPath = @"./Word/";
        //        DirectoryInfo di = new DirectoryInfo(docPath);
        //        Word2PDF w = new Word2PDF();
        //        w.WordTOPDF(di);
        //        sw.Stop();
        //        Console.WriteLine("全部Word已转换为PDF,用时{0}秒", sw.Elapsed);
        //        Console.ReadKey();
        //    }
        //}
    }
}