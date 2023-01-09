using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BookmarksTool.LeiTools.Keygen
{
    public static class DateTimeHelper
    {
        public static void DateTimeLearn()
        {
            //1、DateTime 数字型
            System.DateTime currentTime = new System.DateTime();
            //1.1 取当前年月日时分秒
            currentTime = System.DateTime.Now;
            Console.WriteLine(currentTime);
            //1.2 取当前年
            int 年 = currentTime.Year;
            Console.WriteLine(年);
            //1.3 取当前月
            int 月 = currentTime.Month;
            //1.4 取当前日
            int 日 = currentTime.Day;
            //1.5 取当前时
            int 时 = currentTime.Hour;
            //1.6 取当前分
            int 分 = currentTime.Minute;
            //1.7 取当前秒
            int 秒 = currentTime.Second;
            //1.8 取当前毫秒
            int 毫秒 = currentTime.Millisecond;
        }

        /// <summary>
        /// 判断目前时间是否超过某年，如超过了返回false
        /// </summary>
        /// <param name="expiredYear">过期时间，如判断系统时间是否超过2026年 IsDateExpired(2026) </param>
        /// <returns></returns>
        public static bool IsDateExpired(int expiredYear)
        {
            bool result = false;
            DateTime currentTime = new DateTime();
            currentTime = System.DateTime.Now;
            int year = currentTime.Year;
            //Console.WriteLine("当前时间：" + currentTime);
            //Console.WriteLine("年份：" + year);
            if (year <= expiredYear)
            {
                //Console.WriteLine(year);
                return true;
            }
            return result;
        }
    }
}