using System.IO;

namespace BookmarksTool.LeiTools
{
    class IOHelper
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
    }
}
