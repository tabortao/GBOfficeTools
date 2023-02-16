using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows.Forms;

namespace BookmarksTool.LeiTools.ConfigHelper
{
    public class JSONHelper
    {
        /// <summary>
        /// 读取JSON文件
        /// </summary>
        public static T ReadJSON<T>(string key)
        {
            try
            {
                string path = Application.StartupPath + @"\appsettings.json";
                StreamReader streamReader = new StreamReader(path);
                dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(streamReader.ReadToEnd());
                streamReader.Close();
                return (T)jsonObj[key];
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "/r/n" + e.StackTrace);
            }
            return default;
        }

        /// <summary>
        /// 修改JSON
        /// </summary>
        public static void WriteJSON<T>(string key, T value)
        {
            try
            {
                string path = Application.StartupPath + @"\appsettings.json";
                StreamReader streamReader = new StreamReader(path);
                dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(streamReader.ReadToEnd());
                jsonObj[key] = value;
                streamReader.Close();
                string output = JsonConvert.SerializeObject(jsonObj, Formatting.Indented);
                File.WriteAllText(path, output);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "/r/n" + e.StackTrace);
            }
        }
    }
}