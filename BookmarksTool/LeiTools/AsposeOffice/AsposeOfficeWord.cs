﻿using Aspose.Words;
using System;
using System.Drawing;

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
    }
}