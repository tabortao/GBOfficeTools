using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace BookmarksTool.LeiTools.MSOffice
{
    public class MSWord
    {
        /// <summary>
        /// 判断word文件是否打开
        /// </summary>
        /// <param name="pathName">文件路径</param>
        /// <returns></returns>
        public bool IsFileLocked(string pathName)
        {
            try
            {
                if (!File.Exists(pathName))
                {
                    return false;
                }
                using (var fs = new FileStream(pathName, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    fs.Close();
                }
            }
            catch
            {
                return true;
            }
            return false;
        }

        #region Word转换为PDF

        public static bool Word2PDF(string wordPath)
        {
            //1、打开word
            bool result = false;
            Application application = new Application
            {
                Visible = false
            };
            Document document = null;
            //实现查找路径中word文件,带来筛选，直接选出word文件。
            var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //var wordFiles = Directory.EnumerateFiles(wordPath, "*.doc");
            foreach (var wordFile in wordFiles)
            {
                string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFile); //获取文件名称，不含拓展名。
                string pdfFilePath = wordPath + @"\" + wordFileNameWithoutExtension + ".pdf"; //设置pdf文件存储路径。
                //Console.WriteLine("wordFileNameWithoutExtension" + wordFileNameWithoutExtension);
                //Console.WriteLine("pdfFilePath" + pdfFilePath);

                //循环，转换每一个word文件。
                try
                {
                    if (!File.Exists(@pdfFilePath))//存在PDF，不需要继续转换
                    {
                        document = application.Documents.Open(wordFile);
                        document.ExportAsFixedFormat(pdfFilePath, WdExportFormat.wdExportFormatPDF);
                        //Form1.form1.TextBoxMsg(f.Name + "转换PDF成功!");
                        Console.WriteLine("文件{0}转换PDF成功。", wordFileNameWithoutExtension);
                    }
                    else
                    {
                        File.Delete(@pdfFilePath);
                        document.ExportAsFixedFormat(pdfFilePath, WdExportFormat.wdExportFormatPDF);
                    }
                    result = true;
                }
                catch (Exception e)
                {
                    //System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                    Console.WriteLine(e.Message);
                    result = false;
                }
                finally
                {
                    document.Close();
                }
            }
            //application.Quit();//退出word
            return result;
        }

        /// <summary>
        /// Word转Pdf核心函数
        /// </summary>
        public static bool WordTOPDF(DirectoryInfo docPath)
        {
            bool result = false;
            Application application = new Application
            {
                Visible = false
            };
            Document document = null;
            FileInfo[] files = docPath.GetFiles();

            #region 并行循环

            /*
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
            */

            #endregion 并行循环

            #region 普通循环

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
                            Form1.form1.TextBoxMsg(f.Name + "转换PDF成功!");
                            //Console.WriteLine("文件{0}转换PDF成功", f.Name);
                        }
                        result = true;
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                        //Console.WriteLine(e.Message);
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
                            Form1.form1.TextBoxMsg(f.Name + "转换PDF成功!");
                        }
                        result = true;
                    }
                    catch (Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show("请关闭需要转换的所有word文档。" + "\r\n" + e.Message);
                        result = false;
                    }
                    finally
                    {
                        document.Close();
                    }
                }
            }

            #endregion 普通循环

            return result;
        }

        public static void StartWord2PDF()
        {
            var docPath = Directory.GetCurrentDirectory();
            var path = new DirectoryInfo(docPath);          
            WordTOPDF(path);
            
            
        }

        #endregion Word转换为PDF
    }
}