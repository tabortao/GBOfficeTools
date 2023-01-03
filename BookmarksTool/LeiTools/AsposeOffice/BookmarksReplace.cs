using System;
using System.IO;

namespace BookmarksTool.LeiTools.AsposeOffice
{
    public class BookmarksReplace
    {
        public void StartBookmarksReplace()
        {
            try
            {
                var docPath = Directory.GetCurrentDirectory();
                var path = new DirectoryInfo(docPath);
                var files = Directory.GetFiles(docPath, "*.xlsx", SearchOption.AllDirectories);

                if (files.Length > 1)
                {
                    foreach (var file in files)
                    {
                        //Console.WriteLine("模板：" + Path.GetFileName(file));
                        Form1.form1.TextBoxMsg("模板：" + Path.GetFileName(file));
                    }
                    //Console.WriteLine("存在多个报告模板，文件夹内请保留一个模板，请注意！");
                    Form1.form1.TextBoxMsg("存在多个报告模板，文件夹内请保留一个模板，请注意！");
                    var excelPath = files[0];
                    var excelName = Path.GetFileName(excelPath);
                    //Console.WriteLine("当前计算采用模板为：" + excelName);
                    Form1.form1.TextBoxMsg("当前计算采用模板为：" + excelName);
                    ReportMaker(excelPath, path);
                }
                else if (files.Length == 1)
                {
                    var excelPath = files[0];
                    var excelName = Path.GetFileName(excelPath);
                    //Console.WriteLine("当前计算采用模板为：" + excelName);
                    Form1.form1.TextBoxMsg("当前计算采用模板为：" + excelName);
                    ReportMaker(excelPath, path);
                }
                else
                {
                    //Console.WriteLine("请在当前文件夹中放入Word书签内容批量修改工具.xlsx！");
                    Form1.form1.TextBoxMsg("请在当前文件夹中放入Word书签内容批量修改工具.xlsx！");
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 报告生成器
        /// </summary>
        /// <param name="excelPath">excel文件路径</param>
        /// <param name="path">Word模板路径</param>
        private void ReportMaker(string excelPath, DirectoryInfo path)
        {
            var excel = new LeiTools.AsposeOffice.AsposeOfficeExcel();
            excel.Open(excelPath);
            const int bookmarkNo = 500;//sheet1 书签列书签个数
            var files = path.GetFiles();
            foreach (var f in files)
            {
                if (f.Extension != ".doc" && f.Extension != ".docx") continue;
                try
                {
                    var wordPath = f.FullName;
                    var wordName = f.Name;
                    var word = new LeiTools.AsposeOffice.AsposeOfficeWord();
                    word.Open(wordPath);
                    word.Builder();
                    excel.GetWorksheet(0);//excel转到sheet1
                    for (var i = 0; i <= bookmarkNo; i++)
                    {
                        var bookmarkName = excel.GetCellsValue(i, 3);
                        var bookmarkText = excel.GetCellsValue(i, 2);
                        if (bookmarkName != "0")
                        {
                            //Console.WriteLine(bookmarkName);
                            word.ReplaceBookMark(bookmarkName, bookmarkText);
                        }
                    }

                    if (File.Exists(@"./参评范围.jpg"))
                    {
                        word.ReplaceBookMark("Involved_Range", @"./参评范围.jpg", "IMG");
                    }
                    word.Save(wordPath);
                    //Console.WriteLine(wordName + "  报告生成完成！");
                    Form1.form1.TextBoxMsg(wordName + "  报告生成完成！");
                }
                catch (Exception e)
                {
                    //Console.WriteLine("错误，{0}", e.Message);
                    System.Windows.Forms.MessageBox.Show(e.Message);
                }
            }
        }
    }
}