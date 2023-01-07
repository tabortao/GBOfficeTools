using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace Test
{
    internal class Program
    {
        #region word 转 pdf

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

        #endregion word 转 pdf

        #region 判断文件是否被打开，并测试

        /// <summary>
        /// 判断word文件是否打开
        /// </summary>
        /// <param name="pathName">word文件路径</param>
        /// <returns></returns>
        public static bool IsFileLocked(string pathName)
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

        public static void TestIsFileLocked()
        {
            string path = Directory.GetCurrentDirectory();
            Console.WriteLine(path);
            string pathName = path + @"\08 地下空间利用计算书 居建.docx";
            Console.WriteLine(pathName);
            if (IsFileLocked(pathName))
            {
                Console.WriteLine("文件被打开");
            }
            else
            {
                Console.WriteLine("文件没有被打开");
            }
            Console.ReadLine();
        }

        #endregion 判断文件是否被打开，并测试

        private static void Main(string[] args)
        {
            //TestIsFileLocked();
            string wordPath = @"E:\Code\CSharp\CSharpProjects2023\GBOfficeTools\Test\bin\Debug";
            Word2PDF(wordPath);
            Console.WriteLine();

            Console.ReadLine();
        }
    }
}