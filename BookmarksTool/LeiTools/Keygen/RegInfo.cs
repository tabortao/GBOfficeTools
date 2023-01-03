using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace BookmarksTool.LeiTools.Keygen
{
    /// <summary>
    ///注册信息
    /// </summary>
    public class RegInfo
    {
        /// <summary>
        /// 取本机机器码
        /// </summary>
        public static string GetMachineCode()
        {
            //CPU信息
            string cpuId = DeviceHelper.GetCpuID();
            //磁盘信息
            string diskId = DeviceHelper.GetDiskID();
            //网卡信息
            string MacAddress = DeviceHelper.GetMacAddress();

            string m1 = GetMD5(cpuId + typeof(string).ToString());
            string m2 = GetMD5(diskId + typeof(int).ToString());
            string m3 = GetMD5(MacAddress + typeof(double).ToString());

            string code1 = GetNum(m1, 8);
            string code2 = GetNum(m2, 8);
            string code3 = GetNum(m3, 8);

            return code1 + code2 + code3;
        }

        /// <summary>
        /// 根据机器码产生注册码
        /// </summary>
        /// <param name="machineCode">机器码</param>
        /// <param name="overTime">到期时间</param>
        /// <returns></returns>
        public static string CreateRegisterCode(string machineCode, DateTime overTime)
        {
            int year = int.Parse(overTime.Year.ToString().Substring(2)) + 33;
            int month = overTime.Month + 21;
            int day = overTime.Day + 54;
            int section = machineCode.Length / 4;
            string reg = "";
            int n = 1597;
            for (int i = 0; i < section; i++)
            {
                int sec = int.Parse(machineCode.Substring(i * 4, 4));
                int resu = sec + n;
                if (resu >= 10000)
                {
                    resu = sec - 1597;
                }
                reg += resu;
                n += 1597;
            }
            //插入年月日信息
            reg = InsertNum(reg, year, 0, 8, 4, 6, 7, 1, 3, 2, 5, 9);
            reg = InsertNum(reg, month, 0, 6, 9, 7, 3, 8, 4, 1, 2, 5);
            reg = InsertNum(reg, day, 0, 1, 2, 5, 6, 7, 3, 8, 9, 4);
            return reg.ToString();
        }

        /// <summary>
        /// 检查注册码
        /// </summary>
        /// <param name="registerCode"></param>
        /// <param name="overTime"></param>
        /// <returns></returns>
        public static bool CheckRegister(string registerCode, ref DateTime overTime)
        {
            try
            {
                string machineCode = GetMachineCode();
                //提取年月日
                int day = int.Parse(ExtractNum(ref registerCode, 0, 1, 2, 5, 6, 7, 3, 8, 9, 4));
                int month = int.Parse(ExtractNum(ref registerCode, 0, 6, 9, 7, 3, 8, 4, 1, 2, 5));
                int year = int.Parse(ExtractNum(ref registerCode, 0, 8, 4, 6, 7, 1, 3, 2, 5, 9));
                day -= 54;
                month -= 21;
                year -= 33;
                overTime = new DateTime(year, month, day);
                //核对注册码
                int section = machineCode.Length / 4;
                int n = 1597;
                string reg = "";
                for (int i = 0; i < section; i++)
                {
                    int sec = int.Parse(machineCode.Substring(i * 4, 4));
                    int resu = sec + n;
                    if (resu >= 10000)
                    {
                        resu = sec - 1597;
                    }
                    reg += resu;
                    n = n + 1597;
                }
                return registerCode == reg;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 在指定数字后面插入内容
        /// </summary>
        /// <param name="str"></param>
        /// <param name="num"></param>
        /// <param name="index"></param>
        /// <param name="pmc"></param>
        /// <returns></returns>
        static string InsertNum(string str, int num, int index, params int[] pmc)
        {
            int posi = str.IndexOf(pmc[index].ToString());
            if (posi <= -1)
                return InsertNum(str, num, index + 1, pmc);
            return str.Insert(posi, num.ToString());
        }

        /// <summary>
        /// 提取数字
        /// </summary>
        /// <param name="str"></param>
        /// <param name="index"></param>
        /// <param name="pmc"></param>
        /// <returns></returns>
        static string ExtractNum(ref string str, int index, params int[] pmc)
        {
            int posi = str.IndexOf(pmc[index].ToString());
            if (posi <= -1)
                return ExtractNum(ref str, index + 1, pmc);
            string resu = str.Substring(posi - 2, 2);
            str = str.Remove(posi - 2, 2);
            return resu;
        }

        /// <summary>
        /// 取MD5，32位MD5加密
        /// </summary>
        /// <param name="myString"></param>
        /// <returns></returns>
        public static string GetMD5(string myString)
        {
            MD5 md5 = new MD5CryptoServiceProvider(); //实例化一个md5对像
            //MD5 md5 = MD5.Create();
            byte[] fromData = System.Text.Encoding.Unicode.GetBytes(myString);
            // 加密后是一个字节类型的数组，这里要注意编码UTF8/Unicode等的选择　
            byte[] targetData = md5.ComputeHash(fromData);
            string byte2String = null;
            // 通过使用循环，将字节类型的数组转换为字符串，此字符串是常规字符格式化所得
            for (int i = 0; i < targetData.Length; i++)
            {
                // 将得到的字符串使用十六进制类型格式。格式后的字符是小写的字母，如果使用大写（X）则格式后的字符是大写字符 
                byte2String += targetData[i].ToString("x");
            }
            return byte2String;
        }

        /// <summary>
        /// 取数字
        /// </summary>
        /// <param name="md5"></param>
        /// <param name="len"></param>
        /// <returns></returns>
        static string GetNum(string md5, int len)
        {
            Regex regex = new Regex(@"\d");
            MatchCollection listMatch = regex.Matches(md5);
            string str = "";
            for (int i = 0; i < len; i++)
            {
                str += listMatch[i].Value;
            }
            while (str.Length < len)
            {
                //不足补0
                str += "0";
            }
            return str;
        }
    }
}
