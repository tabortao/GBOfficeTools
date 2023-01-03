using System;
using System.IO;

namespace BookmarksTool.LeiTools.AsposeOffice
{
    public class AsposeOfficeExcel
    {
        private Aspose.Cells.Workbook wb;
        private Aspose.Cells.Worksheet ws;

        /// <summary>
        /// 打开指定路径excel
        /// </summary>
        /// <param name="path">excel 文件路径</param>
        public void Open(string path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                wb = new Aspose.Cells.Workbook(path);
            }
        }

        /// <summary>
        /// 打开空的excel
        /// </summary>
        public void Open()
        {
            wb = new Aspose.Cells.Workbook();
        }

        /// <summary>
        /// 获取工作蒲
        /// </summary>
        /// <param name="sheetIndex">Sheet索引值</param>
        public void GetWorksheet(int sheetIndex)
        {
            ws = wb.Worksheets[sheetIndex];
        }

        public void GetWorksheet(string sheetName)
        {
            ws = wb.Worksheets[sheetName];
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <returns>The cells value.</returns>
        /// <param name="rowIndex">row index 行</param>
        /// <param name="colIndex">col index 列</param>
        public string GetCellsValue(int rowIndex, int colIndex)
        {
            try
            {
                string str = ws.Cells[rowIndex, colIndex].StringValue;
                if (str != "")
                {
                    return str;
                }
                else
                {
                    return "0";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetCellsValue(string cellName)
        {
            try
            {
                string str = ws.Cells[cellName].StringValue;
                if (str != "")
                {
                    return str;
                }
                else
                {
                    return "0";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 单元格写入值
        /// </summary>
        /// <param name="rowIndex">Row index.</param>
        /// <param name="colIndex">Col index.</param>
        /// <param name="value">Value.</param>
        public void SetCellsValue(int rowIndex, int colIndex, string value)
        {
            try
            {
                var cell = ws.Cells[rowIndex, colIndex];
                if (cell.IsMerged)
                {
                    cell.GetMergedRange().PutValue(value, false, false);
                    //Console.WriteLine("写入成功");
                }
                else
                {
                    cell.PutValue(value);
                    //Console.WriteLine("写入成功!");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 把一个数值写入到excel
        /// </summary>
        /// <param name="cellName">单元格名</param>
        /// <param name="value">Value</param>
        public void SetCellsValue(string cellName, string value)
        {
            try
            {
                var cell = ws.Cells[cellName];
                if (cell.IsMerged)
                {
                    cell.GetMergedRange().PutValue(value, false, false);
                    //Console.WriteLine("写入成功");
                }
                else
                {
                    cell.PutValue(value);
                    //Console.WriteLine("写入成功!");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 保存excel
        /// </summary>
        /// <param name="path">Path.</param>
        public void Save(string path)
        {
            wb.Save(path);
        }


    }

}