using System;
using System.IO;

namespace BookmarksTool.LeiTools.AsposeOffice
{
    public class BookmarksReplace
    {
        /// <summary>
        /// 如果没有选择excel文件且没有选择word文件，默认在程序所在文件夹查找并执行
        /// </summary>
        public static void ReportMaker()
        {
            try
            {
                var docPath = Directory.GetCurrentDirectory();
                //var path = new DirectoryInfo(docPath);
                var files = Directory.GetFiles(docPath, "*.xlsx", SearchOption.AllDirectories);

                if (files.Length > 1)
                {
                    foreach (var file in files)
                    {
                        //Console.WriteLine("模板：" + Path.GetFileName(file));
                        MainForm.form1.TextBoxMsg("模板：" + Path.GetFileName(file));
                    }
                    //Console.WriteLine("存在多个报告模板，文件夹内请保留一个模板，请注意！");
                    MainForm.form1.TextBoxMsg("存在多个报告模板，文件夹内请保留一个模板，请注意！");
                    var excelPath = files[0];
                    var excelName = Path.GetFileName(excelPath);
                    //Console.WriteLine("当前计算采用模板为：" + excelName);
                    //Form1.form1.TextBoxMsg("当前计算采用模板为：" + excelName);
                    ReportMaker(excelPath, docPath);
                }
                else if (files.Length == 1)
                {
                    var excelPath = files[0];
                    var excelName = Path.GetFileName(excelPath); // 获取excel文件名，不含路径
                    //Console.WriteLine("当前计算采用模板为：" + excelName);
                    //Form1.form1.TextBoxMsg("当前计算采用模板为：" + excelName);
                    ReportMaker(excelPath, docPath);
                }
                else
                {
                    //Console.WriteLine("请在当前文件夹中放入Word书签内容批量修改工具.xlsx！");
                    MainForm.form1.TextBoxMsg("请在当前文件夹中放入Word书签内容批量修改工具.xlsx！");
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
        public static void ReportMaker(string excelPath, string wordFolderPath)
        {
            MainForm.form1.TextBoxMsg("当前计算采用的Excel模板为：" + Path.GetFileName(excelPath)); //显示excel模板文件名
            var wordFiles = Directory.GetFiles(wordFolderPath, "*.doc"); //从文件夹中，筛选出word
            var excel = new LeiTools.AsposeOffice.AsposeOfficeExcel();
            excel.Open(excelPath);
            const int bookmarkNo = 500;//sheet1 书签列书签个数
            //var files = Directory.GetFiles(docPath, "*.xlsx", SearchOption.AllDirectories);

            foreach (var f in wordFiles)
            {
                //if (f.Extension != ".doc" && f.Extension != ".docx") continue;
                try
                {
                    var word = new LeiTools.AsposeOffice.AsposeOfficeWord();
                    word.Open(f);
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
                    string involvedRangePath = Path.GetDirectoryName(f) + @"\参评范围.jpg";
                    if (File.Exists(involvedRangePath))
                    {
                        word.ReplaceBookMark("Involved_Range", involvedRangePath, "IMG");
                    }
                    word.Save(f);
                    //Console.WriteLine(wordName + "  报告生成完成！");
                    MainForm.form1.TextBoxMsg(f + "  报告生成完成！");
                }
                catch (Exception e)
                {
                    //Console.WriteLine("错误，{0}", e.Message);
                    System.Windows.Forms.MessageBox.Show(e.Message);
                }
            }
        }

        /// <summary>
        /// 选中1个excel文件模板，选择多个word文件，然后批量对书签替换
        /// </summary>
        /// <param name="excelPath">单个excel文件路径</param>
        /// <param name="wordsPath">多个word文件路径，这里是通过选择文件对话框实现，已经过滤了只能选择word文件</param>
        public static void ReportMaker(string excelPath, string[] wordsPath)
        {
            MainForm.form1.TextBoxMsg("当前计算采用的Excel模板为：" + Path.GetFileName(excelPath)); //显示excel模板文件名
            var excel = new LeiTools.AsposeOffice.AsposeOfficeExcel();
            excel.Open(excelPath);
            const int bookmarkNo = 500;//sheet1 书签列书签个数
            //var files = path.GetFiles();
            foreach (var f in wordsPath)
            {
                //if (f.Extension != ".doc" && f.Extension != ".docx") continue;
                try
                {
                    //var wordPath = Path.GetDirectoryName(f); //word文件所在的文件夹路径
                    var wordName = Path.GetFileName(f);
                    var word = new LeiTools.AsposeOffice.AsposeOfficeWord();
                    word.Open(f);
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
                    string involvedRangePath = Path.GetDirectoryName(f)+ @"\参评范围.jpg";
                    if (File.Exists(involvedRangePath))
                    {
                        word.ReplaceBookMark("Involved_Range", involvedRangePath, "IMG");
                    }
                    word.Save(f);
                    //Console.WriteLine(wordName + "  报告生成完成！");
                    MainForm.form1.TextBoxMsg(wordName + "  报告生成完成！");
                }
                catch (Exception e)
                {
                    //Console.WriteLine("错误，{0}", e.Message);
                    System.Windows.Forms.MessageBox.Show(e.Message);
                }
            }
        }

        /// <summary>
        /// 并行实现书签替换
        /// </summary>
        public static void ParallelReportMaker()
        {
            try
            {
                var docPath = Directory.GetCurrentDirectory();
                //var path = new DirectoryInfo(docPath);
                var files = Directory.GetFiles(docPath, "*.xlsx", SearchOption.AllDirectories);

                if (files.Length > 1)
                {
                    foreach (var file in files)
                    {
                        //Console.WriteLine("模板：" + Path.GetFileName(file));
                        MainForm.form1.TextBoxMsg("模板：" + Path.GetFileName(file));
                    }
                    //Console.WriteLine("存在多个报告模板，文件夹内请保留一个模板，请注意！");
                    MainForm.form1.TextBoxMsg("存在多个报告模板，文件夹内请保留一个模板，请注意！");
                    var excelPath = files[0];
                    var excelName = Path.GetFileName(excelPath);
                    //Console.WriteLine("当前计算采用模板为：" + excelName);
                    MainForm.form1.TextBoxMsg("当前计算采用模板为：" + excelName);
                    ParallelReportMaker(excelPath, docPath);
                }
                else if (files.Length == 1)
                {
                    var excelPath = files[0];
                    var excelName = Path.GetFileName(excelPath); // 获取excel文件名，不含路径
                    //Console.WriteLine("当前计算采用模板为：" + excelName);
                    MainForm.form1.TextBoxMsg("当前计算采用模板为：" + excelName);
                    ParallelReportMaker(excelPath, docPath);
                }
                else
                {
                    //Console.WriteLine("请在当前文件夹中放入Word书签内容批量修改工具.xlsx！");
                    MainForm.form1.TextBoxMsg("请在当前文件夹中放入Word书签内容批量修改工具.xlsx！");
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 并行实现书签替换
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="wordFolderPath"></param>
        public static void ParallelReportMaker(string excelPath, string wordFolderPath)
        {
            MainForm.form1.TextBoxMsg("当前计算采用的Excel模板为：" + Path.GetFileName(excelPath)); //显示excel模板文件名
            var wordFiles = Directory.GetFiles(wordFolderPath, "*.doc"); //从文件夹中，筛选出word
            var excel = new LeiTools.AsposeOffice.AsposeOfficeExcel();
            excel.Open(excelPath);
            const int bookmarkNo = 500;//sheet1 书签列书签个数
            //var files = Directory.GetFiles(docPath, "*.xlsx", SearchOption.AllDirectories);

            _ = System.Threading.Tasks.Parallel.ForEach(wordFiles, f =>
              {
                  //if (f.Extension != ".doc" && f.Extension != ".docx") continue;
                  try
                  {
                      var word = new LeiTools.AsposeOffice.AsposeOfficeWord();
                      word.Open(f);
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
                      string involvedRangePath = Path.GetDirectoryName(f) + @"\参评范围.jpg";
                      if (File.Exists(involvedRangePath))
                      {
                          word.ReplaceBookMark("Involved_Range", involvedRangePath, "IMG");
                      }
                      word.Save(f);
                      //Console.WriteLine(wordName + "  报告生成完成！");
                      MainForm.form1.TextBoxMsg(f + "  报告生成完成！");
                  }
                  catch (Exception e)
                  {
                      //Console.WriteLine("错误，{0}", e.Message);
                      System.Windows.Forms.MessageBox.Show(e.Message);
                  }
              });
        }

        /// <summary>
        /// 并行实现书签替换
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="wordsPath"></param>
        public static void ParallelReportMaker(string excelPath, string[] wordsPath)
        {
            MainForm.form1.TextBoxMsg("当前计算采用的Excel模板为：" + Path.GetFileName(excelPath)); //显示excel模板文件名
            var excel = new LeiTools.AsposeOffice.AsposeOfficeExcel();
            excel.Open(excelPath);
            const int bookmarkNo = 500;//sheet1 书签列书签个数
            //var files = path.GetFiles();
            _ = System.Threading.Tasks.Parallel.ForEach(wordsPath, f =>
            {
                //if (f.Extension != ".doc" && f.Extension != ".docx") continue;
                try
                {
                    //var wordPath = Path.GetDirectoryName(f); //word文件所在的文件夹路径
                    var wordName = Path.GetFileName(f);
                    var word = new LeiTools.AsposeOffice.AsposeOfficeWord();
                    word.Open(f);
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
                    string involvedRangePath = Path.GetDirectoryName(f) + @"\参评范围.jpg";
                    if (File.Exists(involvedRangePath))
                    {
                        word.ReplaceBookMark("Involved_Range", involvedRangePath, "IMG");
                    }
                    word.Save(f);
                    //Console.WriteLine(wordName + "  报告生成完成！");
                    MainForm.form1.TextBoxMsg(wordName + "  报告生成完成！");
                }
                catch (Exception e)
                {
                    //Console.WriteLine("错误，{0}", e.Message);
                    System.Windows.Forms.MessageBox.Show(e.Message);
                }
            });
        }
    }
}