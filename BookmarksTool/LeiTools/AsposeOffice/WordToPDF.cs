using Aspose.Words;
using System;
using System.Drawing;
using System.IO;

namespace BookmarksTool.LeiTools.AsposeOffice
{
    class WordToPDF
    {
        #region 批量word转pdf

        /// <summary>
        /// Word2PDF()无参数时，直接将软件所在文件夹内的word转换为PDF。
        /// </summary>
        /// <returns></returns>
        public static void Word2PDF()
        {
            var wordPath = Directory.GetCurrentDirectory();
            DirectoryInfo d = new DirectoryInfo(wordPath);
            FileInfo[] files = d.GetFiles();

            //删除文件夹中~$.doc格式的过程文件
            foreach (var f in files)
            {
                if (f.Name[0] == '~')
                {
                    System.IO.File.Delete(f.FullName);//删除~$.doc格式的过程文件，否则影响程序正常运行。
                }
            }

            //实现查找路径中word文件,带来筛选，直接选出word文件。
            var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //SaveToPDF(wordFiles);
            ParallelSaveToPDF(wordFiles); // 并行循环执行Word转PDF
        }

        /// <summary>
        /// 把某个文件夹内的全部word转换为PDF
        /// </summary>
        /// <param name="wordPath">word文件夹路径</param>
        /// <returns></returns>
        public static void Word2PDF(string wordPath)
        {
            DirectoryInfo d = new DirectoryInfo(wordPath);
            FileInfo[] files = d.GetFiles();
            //删除文件夹中~$.doc格式的过程文件
            foreach (var f in files)
            {
                if (f.Name[0] == '~')
                {
                    System.IO.File.Delete(f.FullName);//删除~$.doc格式的过程文件，否则影响程序正常运行。
                }
            }
            //实现查找路径中word文件,带来筛选，直接选出word文件。
            var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //SaveToPDF(wordFiles);
            ParallelSaveToPDF(wordFiles); // 并行循环执行Word转PDF
        }

        /// <summary>
        /// 选中多个word文件，然后转换为PDF
        /// </summary>
        /// <param name="wordFiles">word文件路径，数组格式</param>
        /// <returns></returns>
        public static void Word2PDF(string[] wordFiles)
        {
            string wordPath = System.IO.Path.GetDirectoryName(wordFiles[0]);
            DirectoryInfo d = new DirectoryInfo(wordPath);
            FileInfo[] files = d.GetFiles();
            //删除文件夹中~$.doc格式的过程文件
            foreach (var f in files)
            {
                if (f.Name[0] == '~')
                {
                    System.IO.File.Delete(f.FullName);//删除~$.doc格式的过程文件，否则影响程序正常运行。
                }
            }
            //实现查找路径中word文件,带来筛选，直接选出word文件。
            //var wordFiles = Directory.GetFiles(wordPath, "*.doc");
            //var wordFiles = Directory.EnumerateFiles(wordPath, "*.doc");
            //SaveToPDF(wordFiles);
            ParallelSaveToPDF(wordFiles); // 并行循环执行Word转PDF
        }

        /// <summary>
        /// 普通循环，批量将Word转PDF
        /// </summary>
        /// <param name="wordFiles"></param>
        /// <returns></returns>
        public static void SaveToPDF(string[] wordFiles)
        {
            //普通循环，转换每一个word文件。
            foreach (var wordFile in wordFiles)
            {
                string wordFileFolder = Path.GetDirectoryName(wordFile);
                string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFile); //获取文件名称，不含拓展名。
                string pdfFilePath = wordFileFolder + @"\" + wordFileNameWithoutExtension + ".pdf"; //设置pdf文件存储路径。

                if (!File.Exists(pdfFilePath))
                {
                    SaveToPDF(wordFile, pdfFilePath);
                }
                else
                {
                    File.Delete(pdfFilePath);
                    SaveToPDF(wordFile, pdfFilePath);
                }
            }
        }

        /// <summary>
        /// 并行循环，批量将Word转PDF
        /// </summary>
        /// <param name="wordFiles"></param>
        /// <returns></returns>
        public static void ParallelSaveToPDF(string[] wordFiles)
        {
            //并行循环，转换每一个word文件。
            System.Threading.Tasks.Parallel.ForEach(wordFiles, wordFile =>
            {
                string wordFileFolder = Path.GetDirectoryName(wordFile);
                string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFile); //获取文件名称，不含拓展名。
                string pdfFilePath = wordFileFolder + @"\" + wordFileNameWithoutExtension + ".pdf"; //设置pdf文件存储路径。
                if (!File.Exists(pdfFilePath))
                {
                    SaveToPDF(wordFile, pdfFilePath);
                }
                else
                {
                    File.Delete(pdfFilePath);
                    SaveToPDF(wordFile, pdfFilePath);
                }
            });
        }

        /// <summary>
        /// 单个word文件转换为pdf
        /// </summary>
        /// <param name="wordFilePath">word文件路径，含文件名及拓展名</param>
        /// <param name="pdfFilePath">pdf文件路径，含文件名及拓展名</param>
        public static bool SaveToPDF(string wordFilePath, string pdfFilePath)
        {
            bool result;
            string wordFileNameWithoutExtension = Path.GetFileNameWithoutExtension(wordFilePath); //获取文件名称，不含拓展名。
            if (LeiTools.IOHelper.IsFileLocked(wordFilePath))  //判断文件是否被锁定
            {
                MainForm.form1.TextBoxMsg2(wordFileNameWithoutExtension + "转PDF失败，请关闭Word文件后重试！！!");
                return false;
            }
            else
            {
                try
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(wordFilePath);
                    doc.Save(pdfFilePath, Aspose.Words.SaveFormat.Pdf);
                    MainForm.form1.TextBoxMsg2(wordFileNameWithoutExtension + "，转PDF成功!");
                    return true;
                }
                catch (Exception e)
                {
                    MainForm.form1.TextBoxMsg2(e.Message);
                    result = false;
                }
            }
            return result;
        }

        #endregion 批量word转pdf
    }
}
