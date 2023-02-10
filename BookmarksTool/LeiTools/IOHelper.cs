using System;
using System.IO;
using System.Windows.Forms;

namespace BookmarksTool.LeiTools
{
    internal class IOHelper
    {
        /// <summary>
        /// 判断文件是否被锁定
        /// </summary>
        /// <param name="pathName"></param>
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

        /// <summary>
        /// 按照数组，创建文件夹
        /// </summary>
        /// <param name="path"></param>
        public static void Mkdirs(string[] path)
        {
            foreach (var f in path)
            {
                try
                {
                    if (!Directory.Exists(f))
                    {
                        Directory.CreateDirectory(f);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Exception: " + e.Message);
                }
            }
        }

        /// <summary>
        /// 根据ini配置文件，批量创建文件夹
        /// </summary>
        /// <param name="iniFilePath">ini文件夹配置文件路径</param>
        /// <param name="folderPath">要创建的文件夹路径</param>
        public static void BatchCreateFolder(string folderPath)
        {
            //读取ini文件
            //string iniFilePath = folderPath + "CreateFolder.ini";
            string iniFilePath = System.IO.Directory.GetCurrentDirectory() + @"\Data\CreateFolder.ini";
            string[] DirNameLists = File.ReadAllLines(iniFilePath);

            // 创建文件夹

            foreach (string item in DirNameLists)
            {
                string FilePath = folderPath + @"\" + item;
                try
                {
                    if (!Directory.Exists(FilePath))
                    {
                        Directory.CreateDirectory(FilePath);
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Exception: " + e.Message);
                }
            }
        }
    }
}