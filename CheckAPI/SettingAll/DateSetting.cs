using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Configuration;
using Nager.Date;
using Nager.Date.Model;
using Newtonsoft.Json;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using CheckAPI.Models;
using CheckAPI.Controllers;

namespace CheckAPI.SettingAll
{
    public class DateSetting
    {
        #region 共用

        public static bool TodayYOne(string TxtDate)
        {
            DateTime dateTime = SendYMDController.DateString(TxtDate);
            // 使用 DateTime 的屬性 Year 獲取今年的年份
            int year = dateTime.Year;

            // 通過將年份設置為 1 月 1 日來獲取今年的第一天
            DateTime firstDayOfYear = new DateTime(year, 1, 1);
            if (dateTime.ToString("yyyy/MM/dd") == firstDayOfYear.ToString("yyyy/MM/dd")) {
                return true;
            }
            else { 
                return false;
            }

        }

        public static bool TodayMOne(string TxtDate)
        {
            // 獲取當前日期和時間
            DateTime dateTime = SendYMDController.DateString(TxtDate);

            // 使用 DateTime 的屬性 Year 和 Month 獲取當前的年份和月份
            int year = dateTime.Year;
            int month = dateTime.Month;

            // 通過將當前年份和月份設置為 1 日來獲取當前月的第一天
            DateTime firstDayOfMonth = new DateTime(year, month, 1);
            if (dateTime.ToString("yyyy/MM/dd") == firstDayOfMonth.ToString("yyyy/MM/dd"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }




        /// <summary>
        /// 本月最後一天平日
        /// </summary>
        /// <param name="lastDayOfMonth"></param>
        /// <returns></returns>
        private static DateTime FindLastWeekday(DateTime lastDayOfMonth)
        {
            while (lastDayOfMonth.DayOfWeek == DayOfWeek.Saturday || lastDayOfMonth.DayOfWeek == DayOfWeek.Sunday)
            {
                lastDayOfMonth = lastDayOfMonth.AddDays(-1);
            }
            return lastDayOfMonth;
        }
        //判斷最後一天

        public static DateTime LastDayOfMonth(DateTime lastDayOfMonth) {
            string Weekly = lastDayOfMonth.DayOfWeek.ToString("d");//星期幾
            if (DateSetting.HolidayTest(lastDayOfMonth) != "" || Weekly == "6" || Weekly == "0") {
                return LastDayOfMonth(lastDayOfMonth.AddDays(-1) );
            }
            else {
                return lastDayOfMonth;
            }
        }
        /// <summary>
        /// 判斷最後一天是幾日
        /// </summary>
        /// <param name="tomorrow"></param>
        /// <param name="tomorrowOld"></param>
        /// <returns></returns>
        private static bool holidayBool(DateTime tomorrow, DateTime tomorrowOld) {
            //今天日期
           // DateTime today = DateTime.Now;
           // DateTime today = tomorrowOld; //DateTime.Now;
            //最後一天
            DateTime lastDayOfMonth = new DateTime(tomorrowOld.Year, tomorrowOld.Month, DateTime.DaysInMonth(tomorrowOld.Year, tomorrowOld.Month));
            lastDayOfMonth = LastDayOfMonth(lastDayOfMonth);//扣掉資料庫抓到的值，所得到的最後一天，包含扣六日
            int maxDay = lastDayOfMonth.Day;//最後一天天數
            int dayOfMonth = tomorrowOld.Day; // 獲取今天是這個月的第幾天

            //明天
            tomorrow = tomorrow.AddDays(1);
            string Weekly = tomorrow.DayOfWeek.ToString("d");//星期幾
            string WeeklyOld = tomorrowOld.DayOfWeek.ToString("d");//今天星期幾
           // if (DateSetting.HolidayTest(tomorrowOld) != "" || Weekly == "6" || Weekly == "0") return false;

            int maxOldDay = 7 - int.Parse(WeeklyOld);//這禮拜剩下幾天
            //今天不是最後一天
            if (maxDay == dayOfMonth)
            {
                return true;//今天為最後一天
            }
            else {
                return false;
            }           
        }

        /// <summary>
        /// 判斷最後一天年是幾日
        /// </summary>
        /// <param name="tomorrow"></param>
        /// <param name="tomorrowOld"></param>
        /// <returns></returns>
        private static bool holidayYBool(DateTime tomorrow, DateTime tomorrowOld)
        {
            //今天日期
            // DateTime today = DateTime.Now;
            // DateTime today = tomorrowOld; //DateTime.Now;
            //今年最後一天
            DateTime lastDayOfMonth = new DateTime(tomorrowOld.Year, 12, DateTime.DaysInMonth(tomorrowOld.Year, 12));
            lastDayOfMonth = LastDayOfMonth(lastDayOfMonth);//扣掉資料庫抓到的值，所得到的最後一天，包含扣六日
            int maxDay = lastDayOfMonth.Day;//最後一天天數
            int dayOfMonth = tomorrowOld.Day; // 獲取今天是這個月的第幾天

            //明天
            tomorrow = tomorrow.AddDays(1);
            string Weekly = tomorrow.DayOfWeek.ToString("d");//星期幾
            string WeeklyOld = tomorrowOld.DayOfWeek.ToString("d");//今天星期幾
                                                                   // if (DateSetting.HolidayTest(tomorrowOld) != "" || Weekly == "6" || Weekly == "0") return false;

            int maxOldDay = 7 - int.Parse(WeeklyOld);//這禮拜剩下幾天
            //今天不是今年最後一天
            if (maxDay == dayOfMonth)
            {
                return true;//今天為最後一天
            }
            else
            {
                return false;
            }
        }

        #endregion
        #region 判斷是否資料庫假日
        //假日測試
        public static string HolidayTest(DateTime dateOne)
        {
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                select Holiday 
                from HolidayList 
                where Holiday=@Holiday
                "
            };
            cmd.Parameters.Add("@Holiday", SqlDbType.Date).Value = dateOne;
            string Result = GetSQLScalar(cmd, Sqlconn);
            return Result;
        }
        //假日測試
        public static string HolidayLike(string datelike)
        {
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                select Holiday 
                from HolidayList 
                where Holiday like @Holidaylike
                "
            };
            cmd.Parameters.Add("@Holidaylike", SqlDbType.Variant).Value = datelike;
            string Result = GetSQLScalar(cmd, Sqlconn);
            return Result;
        }
        #endregion

        #region  平日最後一天，是否要出單 
        /// <summary>
        /// 平日月最後一天，是否要出單
        /// </summary>
        public static bool FinalDay(string TxtDate)
        {
            // 獲取當前日期
            DateTime today = SendYMDController.DateString(TxtDate);//DateTime.Now;
            string Weekly = today.DayOfWeek.ToString("d");//星期幾
            //特別人日、六、日
            if (DateSetting.HolidayTest(DateTime.Now) != "" || Weekly == "6" || Weekly == "0") return false;
            // 找到本月的最後一天
            DateTime lastDayOfMonth = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month));

            //明天日期
            DateTime tomorrow = today.AddDays(1);
            // 获取本月的第一天
            DateTime firstDayOfMonth = new DateTime(today.Year, today.Month, 1);
            //判斷是否為最後一天假日
            bool OKdateTime= holidayBool(today,today);
            return OKdateTime;
        }
        /// <summary>
        /// 平日年最後一天，是否要出單
        /// </summary>
        public static bool FinalYDay(string TxtDate)
        {
            // 獲取當前日期
            DateTime today = SendYMDController.DateString(TxtDate);//DateTime.Now;
            string Weekly = today.DayOfWeek.ToString("d");//星期幾
            //特別人日、六、日
            if (DateSetting.HolidayTest(DateTime.Now) != "" || Weekly == "6" || Weekly == "0") return false;
            // 找到本年的最後一天
            DateTime lastDayOfMonth = new DateTime(today.Year, 12, DateTime.DaysInMonth(today.Year, 12));
            //明天日期
            DateTime tomorrow = today.AddDays(1);
            // 本年的第一天
            DateTime firstDayOfMonth = new DateTime(today.Year, 1, 1);
            //判斷是否為最後一天假日
            bool OKdateTime = holidayYBool(today, today);
            return OKdateTime;
        }

        #endregion

        #region SQLCmd
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        private static string GetSQLScalar(SqlCommand cmd, string Sqlconn)
        {
            string Result = "";
            using (SqlConnection conn = new SqlConnection(Sqlconn))
            {
                try
                {
                    cmd.Connection = conn;
                    conn.Open();
                    object oo = cmd.ExecuteScalar();
                    if (oo != null)
                        Result = oo.ToString();
                    conn.Close();
                    cmd.Dispose();
                }
                catch
                {
                    Result = "";
                }
                return Result;
            }
        }
        private static string GetSQLNonQuery(SqlCommand cmd, string Sqlconn)
        {
            string Result = "ok";
            try
            {
                using (SqlConnection conn = new SqlConnection(Sqlconn))
                {
                    cmd.Connection = conn;
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    cmd.Dispose();
                }
            }
            catch (Exception ex)
            {
                string e = ex.ToString();
                Result = "";
            }
            return Result;
        }
        private static DataTable GetSQLDataTable(SqlCommand cmd, string Sqlconn)
        {
            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection conn = new SqlConnection(Sqlconn))
                {
                    cmd.Connection = conn;
                    conn.Open();
                    SqlDataReader sdr = cmd.ExecuteReader();
                    dt.Load(sdr);
                    conn.Close();
                    cmd.Dispose();
                }
            }
            catch (Exception ex)
            {
                string e = ex.ToString();
            }
            return dt;
        }
        private void SaveLog(string json)
        {
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                insert into Json_Log 
                ( sendjson, create_at) 
                values 
                (@sendjson,@create_at) 
                ")
            };
            cmd.Parameters.Add("@sendjson", SqlDbType.VarChar).Value = json;
            cmd.Parameters.Add("@create_at", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            GetSQLNonQuery(cmd, Sqlconn);
        }
        #endregion



    }
}
