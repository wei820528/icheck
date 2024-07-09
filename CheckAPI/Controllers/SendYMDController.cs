using CheckAPI.SettingAll;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace CheckAPI.Controllers
{
    public class SendYMDController : Controller
    {
        // GET: SendYMD
        public ActionResult Index()
        {
            return View();
        }
        #region 判斷產生
        /// <summary>
        /////每週星期幾的List 
        /// </summary>
        /// <param name="Weekly"></param>
        /// <param name="FAB"></param>
        /// <returns></returns>
        public static DataTable GetWeeklyList(string Weekly, string FAB)
        {
            SqlCommand cmd = new SqlCommand { };
            string FABWhere = "";
            if (FAB != "")
            {
                FABWhere = " and FAB=@FAB";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }
            cmd.CommandText = string.Format(@"
                select FAB,TableID,TableName 
                from Tables
                where TableEnable='1'
                and WeeklyCycle like @Weekly {0}
                ", FABWhere);
            cmd.Parameters.Add("@Weekly", SqlDbType.VarChar).Value = "%" + Weekly + "%";
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            return dt;
        }
        /// <summary>
        /// 設定日期的
        /// </summary>
        /// <param name="Daily"></param>
        /// <returns></returns>
        //每月幾號的List
        public static DataTable GetDailyList(string Daily, string FAB)
        {
            if (Daily.Length < 2) Daily = "0" + Daily;
            SqlCommand cmd = new SqlCommand { };
            string FABWhere = "";
            if (FAB != "")
            {
                FABWhere = " and FAB=@FAB";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }
            cmd.CommandText = string.Format(@"
                select FAB,TableID,TableName 
                from Tables
                where TableEnable='1'
                and MonthCycle like @Daily {0}
                ", FABWhere);

            cmd.Parameters.Add("@Daily", SqlDbType.VarChar).Value = "%" + Daily + "%";
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            return dt;
        }
        /// <summary>
        /// 找尋月表單
        /// </summary>
        /// <param name="Month"></param>
        /// <param name="FAB"></param>
        /// <returns></returns>
        public static DataTable GetYearList(string Month, string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            if (Month.Length < 2) Month = "0" + Month;
            cmd.Parameters.Add("@Month", SqlDbType.VarChar).Value = "%" + Month + "%";
            string FABWhere = "";
            if (FAB != "")
            {
                FABWhere = " and FAB=@FAB";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }

            cmd.CommandText = string.Format(@"
                select FAB,TableID,TableName 
                from Tables
                where TableEnable='1'
                and YearCycle like @Month {0}
                ", FABWhere);

            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            return dt;
        }
        /// <summary>
        /// 開始時間
        /// </summary>
        /// <param name="dateString"></param>
        /// <returns></returns>
        public static DateTime DateString(string dateString)
        {
            // 獲取當前日期
            DateTime date = DateTime.Now;
            if (dateString != "")
            {
                if (DateTime.TryParse(dateString, out DateTime parsedDate)) return DateTime.Parse(dateString);
            }
            return date;
        }
        /// <summary>
        /// 結束時間
        /// </summary>
        /// <param name="dateString"></param>
        /// <returns></returns>
        public static DateTime DateStringEnd(string dateString)
        {
            // 獲取當前日期
            DateTime today = DateTime.Now;
            DateTime date = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month)); ;
            if (dateString != "")
            {
                if (DateTime.TryParse(dateString, out DateTime parsedDate)) return DateTime.Parse(dateString);             
            }
            return date;
        }

        /// <summary>
        /// 所有這個月表單，簡要
        /// </summary>
        /// <param name="date"></param>
        /// <param name="FAB"></param>
        /// <returns></returns>
        public static DataTable GetAllMForm(DateTime date, string FAB)
        {
            string Month = date.ToString("yyyyMM");
            string Day = date.ToString("yyyyMMdd");
            SqlCommand cmd = new SqlCommand();
            if (Month.Length < 2) Month = "0" + Month;
            cmd.Parameters.Add("@Day", SqlDbType.VarChar).Value = "%" + Day; //"%" + Month + "[0-3][0-9]";// "%" + Month + "%";
            cmd.Parameters.Add("@Month", SqlDbType.VarChar).Value = "%" + Month + "M";
            string FABWhere = "";
            if (FAB != "")
            {
                FABWhere = " and a.FAB=@FAB";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }
            cmd.CommandText = string.Format(@"
                 SELECT a.*,b.TableName,b.TableType,b.UserID as BOSSID,c.UserName,c.UserMail,c.IsBoss,d.SerialNumber
                  FROM Datas as a left join Tables as b on a.TableID=b.TableID
                  left join Users as c on a.UserID=c.UserID 
                  left join TablesType as d on b.TableType=d.TableType
                  where ( a.Doc like @Day or  a.Doc like @Month ) {0} 
                  order by a.FAB desc,a.AliveTime desc
                ", FABWhere);

            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            return dt;
        }
        /// <summary>
        /// 所有這個月表單，詳細
        /// </summary>
        /// <param name="date"></param>
        /// <param name="FAB"></param>
        /// <returns></returns>
        public static DataTable GetAllMForm_List(DateTime date, string FAB)
        { 
            string Month = date.ToString("yyyyMM");
            string Day = date.ToString("yyyyMMdd");
            SqlCommand cmd = new SqlCommand();
            if (Month.Length < 2) Month = "0" + Month;
            cmd.Parameters.Add("@Day", SqlDbType.VarChar).Value = "%" + Day; //"%" + Month + "[0-3][0-9]";// "%" + Month + "%";
            cmd.Parameters.Add("@Month", SqlDbType.VarChar).Value = "%" + Month + "M";
            string FABWhere = "";
            if (FAB != "")
            {
                FABWhere = " and a.FAB=@FAB";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }
            cmd.CommandText = string.Format(@"
            SELECT a.*,
            b.TableEnable,b.TableName,b.TableType,b.UserID as bossID,
            e.UserName,e.UserMail,e.IsAdmin,f.SerialNumber,
            c.ItemID,c.ItemValue,
            d.ItemName,d.ItemType,d.ItemMin,d.ItemMax,d.ItemType,d.ItemSort,d.ItemContent
            FROM Datas as a 
            left join Tables as b on a.TableID=b.TableID
            left join DatasItem as c  on a.Doc= c.Doc
            left join TablesItem as d on d.ItemID= c.ItemID and d.TableID=a.TableID
            left join Users as e on a.UserID=e.UserID
            left join TablesType as f on b.TableType=f.TableType
            where ( a.Doc like @Day or  a.Doc like @Month ) {0} 
            order by a.Doc desc,d.ItemSort asc
                ", FABWhere);

            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            return dt;
        }


        #endregion

        #region 日排成
        public static string C_D_CSV_ALL(string TxtDate)
        {
            //判斷是否平日月產生 
            SqlCommand cmd = new SqlCommand();
            string ok = "";
            //先取得廠區 
            try
            {
                //取得廠區
                cmd.CommandText = @"select FAB from Factories";
                DataTable FABAll = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
                if (FABAll.Rows.Count != 0)
                {
                    foreach (DataRow FABOne in FABAll.Rows)
                    {
                        //搜尋這個月表單
                        ok += "廠區:" + FABOne["FAB"].ToString() + "月表單";
                        ok += D_CSV(TxtDate, FABOne["FAB"].ToString());
                        ok += "廠區:" + FABOne["FAB"].ToString() + "月詳細表單";
                        ok += D_List_CSV(TxtDate, FABOne["FAB"].ToString());
                    }
                }
                return ok;
            }
            catch 
            {

                return ok;
            }
        }

        /// <summary>
        /// 日排成表單明細
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string D_List_CSV(string dateString, string FAB)
        {
            // 獲取當前日期
            DateTime date = DateString(dateString);
            // 這個月的第一天
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);

            // 這個月的最後一天
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = date;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = firstDayOfMonth;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = lastDayOfMonth;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            //正常產生 08C1_20240112.csv
            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
            //產生csv
            string FilePath = "";
            string url = "";
            string Language = "UTF-8";
            StringBuilder lines = new StringBuilder();

            lines.AppendLine(line);
            bool okData = Download_CSV(FilePath, url, lines, Language);
            return line;
        }

        /// <summary>
        /// 日排成表單明細
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string D_CSV(string dateString, string FAB)
        {
            // 獲取當前日期
            DateTime date = DateString(dateString);
            // 這個月的第一天
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);

            // 這個月的最後一天
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = date;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = firstDayOfMonth;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = lastDayOfMonth;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""規格表單編號""";
            string FilePath = "";
            string url = "";
            string Language = "UTF-8";
            StringBuilder lines = new StringBuilder();
            lines.AppendLine(line);
            bool okData = Download_CSV(FilePath, url, lines, Language);
            return line;

        }

        #endregion
        public static bool Download_CSV(string FilePath, string url, StringBuilder lines, string Language)
        {
            bool ok = false;
            // StringBuilder lines = new StringBuilder();
            if (!Directory.Exists(url)) Directory.CreateDirectory(url);
            //*存文字
            try
            {
                System.IO.File.WriteAllText(FilePath, lines.ToString(), Encoding.GetEncoding(Language));
                ok = true;
            }
            catch 
            {
                ok = false;
            }
            return ok;
        }
    }
}