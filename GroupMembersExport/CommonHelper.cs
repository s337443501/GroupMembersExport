using Newtonsoft.Json.Linq;
using SufeiUtil;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace GroupMembersExport
{
    public class CommonHelper
    {
        public static string GetHtml(string url, string cookies = "")
        {
            try
            {
                HttpHelper http = new HttpHelper();
                HttpItem item = new HttpItem()
                {
                    URL = url,
                    Method = "get",
                    Cookie = cookies,
                    UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36",
                    Accept = "text/html, application/xhtml+xml, */*",
                    ContentType = "text/html",
                    Allowautoredirect = true,
                };
                HttpResult result = http.GetHtml(item);
                return result.Html;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        public static string GetCookieValue(string Cookies, string CookieName)
        {
            try
            {
                foreach (string cookie in Cookies.Split(';'))
                {
                    string Name = cookie.Split('=')[0].Trim();
                    string Value = cookie.Split('=')[1].Trim();
                    if (Name == CookieName)
                    {
                        return Value;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        public static string Getbkn(string skey)
        {
            try
            {
                string p_skey = skey;
                long hash = 5381;
                for (int i = 0; i < p_skey.Length; i++)
                {
                    hash += (hash << 5) + p_skey[i];
                }
                long g_tk = hash & 0x7fffffff;
                return g_tk.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

        public static void ExportExcel(DataGridView dgv)
        {
            try
            {
                if (dgv.RowCount > 0)
                {
                    ExcelHelper.ExportExcel(dgv);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// 时间转换 - 数字转时间
        /// </summary>
        /// <param name="originalTime"></param>
        /// <returns></returns>
        public static string NumToTime(string originalTime)
        {
            try
            {
                long unixTimeStamp = Convert.ToInt64(originalTime);
                DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1)); // 当地时区
                DateTime dt = startTime.AddSeconds(unixTimeStamp);
                return dt.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }
    }
}
