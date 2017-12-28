using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    public class StringHelper
    {
        /// <summary>
        /// 截取字符串函数
        /// </summary>
        /// <param name="str">所要截取的字符串</param>
        /// <param name="num">截取字符串的长度</param>
        /// <param name="flg">true:加...,flase:不加</param>
        /// <returns></returns>
        public static string GetSubString(string str, int num, bool flg)
        {
            #region
            ASCIIEncoding ascii = new ASCIIEncoding();
            int tempLen = 0;
            string tempString = "";
            byte[] s = ascii.GetBytes(str);
            for (int i = 0; i < s.Length; i++)
            {
                if ((int)s[i] == 63)
                {
                    tempLen += 2;
                }
                else
                {
                    tempLen += 1;
                }

                try
                {
                    tempString += str.Substring(i, 1);
                }
                catch
                {
                    break;
                }

                if (tempLen > num)
                    break;
            }
            //如果截过则加上半个省略号
            byte[] mybyte = System.Text.Encoding.Default.GetBytes(str);
            if (mybyte.Length > num)
                if (flg)
                {
                    tempString += "...";
                }
            return tempString;
            #endregion
        }
    }
}
