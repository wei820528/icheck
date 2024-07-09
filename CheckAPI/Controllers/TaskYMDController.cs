using CheckAPI.Models;
using CheckAPI.SettingAll;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Http.Results;
using System.Web.Mvc;

namespace CheckAPI.Controllers
{
    public class TaskYMDController : Controller
    {
        #region 月+日排成
        //手動
        public  string C_DM_CSV_ALL(string TxtDate)
        {
            string ok = C_M_CSV_ALL(TxtDate);
            ok += SendYMDController.C_D_CSV_ALL(TxtDate);
            return ok;
        }
        //自動
        public  string C_DM_Automatic_CSV_ALL()
        {
            DateTime today = DateTime.Now;
            bool OKholiday = DateSetting.FinalDay("");
            string ok = "";
            if (OKholiday)
            {
                ok += "今天是這個月最後一天平日，";
                ok += C_M_CSV_ALL(today.ToString("G"));//執行月CSV產生
               
            }
            else
            {
                ok = "今天不是這個月最後一天平日，";
            }
            string Weekly = today.DayOfWeek.ToString("d");//星期2 
            if (DateSetting.HolidayTest(DateTime.Now) != "" || Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly + ",週未假日不出單";
            else
            {
                //執行日排成產生
                ok += "<br>" + SendYMDController.C_D_CSV_ALL(today.ToString("G"));
            }
            return ok;
        }
        #endregion

        #region 月排成
        /// <summary>
        /// 同C_M_CSV_ALL
        /// </summary>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string C_M_CSV_ALL(string TxtDate)
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
                        ok += M_CSV(TxtDate, FABOne["FAB"].ToString());
                        ok += "廠區:" + FABOne["FAB"].ToString() + "月詳細表單";
                        ok += M_List_CSV(TxtDate, FABOne["FAB"].ToString());
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
        /// 同C_M_CSV_ALL
        /// </summary>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public  string C_M_CSV_ALL_Url(string TxtDate)
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
                        ok += M_CSV(TxtDate, FABOne["FAB"].ToString());
                        ok += "廠區:" + FABOne["FAB"].ToString() + "月詳細表單";
                        ok += M_List_CSV(TxtDate, FABOne["FAB"].ToString());
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
        /// 月排成表單
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string M_CSV(string dateString, string FAB)
        {
            // 獲取當前日期
            DateTime date = SendYMDController.DateString(dateString);
            // 這個月的第一天
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);

            // 這個月的最後一天
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = date;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = firstDayOfMonth.ToString("yyyyMMdd") + " 00:00:00"; ;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = lastDayOfMonth.ToString("yyyyMMdd") + " 23:59:59"; ;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            //正常產生 08C1_20240112.csv
            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
            //產生csv            
            string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + date.ToString("yyyyMMdd") + ".csv";
            string url = ConfigurationManager.AppSettings["TxtUrl"];
            string Language = "UTF-8";
            StringBuilder lines = new StringBuilder();
            StringBuilder Result = new StringBuilder();
            DataTable dt = SendYMDController.GetAllMForm(date, FAB);
            lines.AppendLine(line);
            if (dt.Rows.Count != 0)
            {
                line += "<br/>";
                Result.AppendLine(line);
                //*
                foreach (DataRow dr in dt.Rows)
                {
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";
                    if (dr["IsFinishedTime"].ToString() != "")
                    {
                        line += @"=""" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @""",";
                    }
                    else
                    {
                        line += @"="""",";
                    }

                    line += @"=""" + dr["UserName"].ToString() + @""",";
                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";
                    lines.AppendLine(line);
                    line += "<br/>";
                    Result.AppendLine(line);
                }
                bool okData = SendYMDController.Download_CSV(FilePath, url, lines, Language);
            }

           
            return line;
        }

        /// <summary>
        /// 月排成表單明細
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string M_List_CSV(string dateString, string FAB)
        {
            string Today = "", TodayF = "", TodayL = "";
            // 獲取當前日期
            DateTime date = SendYMDController.DateString(dateString);
            // 這個月的第一天
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            // 這個月的最後一天
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            SqlCommand cmd = new SqlCommand();
            Today = "%" + date.ToString("yyyyMMdd");
            TodayF = date.ToString("yyyyMMdd") + " 00:00:00";
            TodayL = date.ToString("yyyyMMdd") + " 23:59:59";
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = Today;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = TodayF;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = TodayL;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;

            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""項目說明"",=""規格表單編號""";
            string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + date.ToString("yyyyMMdd") + "list.csv";
            string url = ConfigurationManager.AppSettings["TxtUrl"];
            string Language = "UTF-8";
            DataTable dt = SendYMDController.GetAllMForm_List(date, FAB);
            StringBuilder lines = new StringBuilder();
            StringBuilder Result = new StringBuilder();
            lines.AppendLine(line);
            if (dt.Rows.Count != 0)
            {

                line += "<br/>";
                Result.AppendLine(line);
                //*
                foreach (DataRow dr in dt.Rows)
                {
                    string ItemTypeTxt = "正常／異常類別";
                    string ItemNameTxt = dr["ItemName"].ToString();
                    string IsError = "N";
                    string ErrorMsg = "";
                    if (dr["ItemType"].ToString() == "1")
                    {
                        if (dr["ItemValue"].ToString() != "OK")
                        {
                            IsError = "Y";
                            ErrorMsg = dr["ItemValue"].ToString();
                        }
                    }
                    if (dr["ItemType"].ToString() == "4")
                    {
                        if (dr["ItemValue"].ToString() != "N/A")
                        {
                            IsError = "Y";
                            ErrorMsg = dr["ItemValue"].ToString();
                        }
                    }
                    if (dr["ItemType"].ToString() == "2")
                    {
                        ItemTypeTxt = "數字類別";
                        ItemNameTxt = ItemNameTxt + "(" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + ")";
                        try
                        {
                            float Min = float.Parse(dr["ItemMin"].ToString());
                            float Max = float.Parse(dr["ItemMax"].ToString());
                            float Value = float.Parse(dr["ItemValue"].ToString());
                            if (Value < Min)
                            {
                                IsError = "Y";
                                ErrorMsg = dr["ItemValue"].ToString() + "低於" + dr["ItemMin"].ToString() + " (" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + ")";
                            }
                            else if (Value > Max)
                            {
                                IsError = "Y";
                                ErrorMsg = dr["ItemValue"].ToString() + "高於" + dr["ItemMax"].ToString() + " (" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + ")";
                            }
                        }
                        catch
                        {
                            IsError = "Y";
                            ErrorMsg = "資料異常";
                        }
                    }
                    if (dr["ItemType"].ToString() == "3") ItemTypeTxt = "文字類別";

                    //"廠別,巡檢單號,巡檢表單名稱,巡檢時間,人員工號,巡檢人員,項目名稱,巡檢紀錄,是否異常,異常原因";
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";

                    if (dr["IsFinishedTime"].ToString() != "")
                    {
                        line += @"=""" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @""",";
                    }
                    else
                    {
                        line += @"="""",";
                    }

                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["UserName"].ToString() + @""",";
                    line += @"=""" + ItemNameTxt + @""",";//項目名稱
                                                          // line += @"=""" + dr["SerialNumber"].ToString() + @""",";
                    /*/
                    if (dr["ItemType"].ToString() == "2")//標準值
                        line += @"=""" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + @""",";
                    else
                        line += @"=""" + "正常" + @""",";//正常
                    //*/
                    if (dr["ItemValue"].ToString() == "OK")//巡檢紀錄
                        line += @"=""" + "正常" + @""",";
                    else
                        line += @"=""" + dr["ItemValue"].ToString() + @""",";

                    //line += @"""" + ItemTypeTxt + @""",";
                    line += @"=""" + IsError + @""",";//是否異常
                    line += @"=""" + ErrorMsg + @""",";//異常原因
                    line += @"=""" + dr["ItemContent"].ToString() + @""",";//項目說明
                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";//規格表單編號

                    lines.AppendLine(line);
                    line += "<br/>";
                    Result.AppendLine(line);
                }
                bool okData = SendYMDController.Download_CSV(FilePath, url, lines, Language);
            }
          
            return line;
        }


        #endregion
        #region 日排成
        #endregion



        #region 年排成
        /// <summary>
        /// 同C_M_CSV_ALL
        /// </summary>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string C_Y_CSV_ALL(string TxtDate)
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
                        ok += M_CSV(TxtDate, FABOne["FAB"].ToString());
                        ok += "廠區:" + FABOne["FAB"].ToString() + "月詳細表單";
                        ok += M_List_CSV(TxtDate, FABOne["FAB"].ToString());
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
        /// 月排成表單
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string Y_CSV(string dateString, string FAB)
        {
            // 獲取當前日期
            DateTime date = SendYMDController.DateString(dateString);
            // 這個月的第一天
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);

            // 這個月的最後一天
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = date;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = firstDayOfMonth.ToString("yyyyMMdd") + " 00:00:00"; ;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = lastDayOfMonth.ToString("yyyyMMdd") + " 23:59:59"; ;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            //正常產生 08C1_20240112.csv
            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
            //產生csv            
            string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + date.ToString("yyyyMMdd") + ".csv";
            string url = ConfigurationManager.AppSettings["TxtUrl"];
            string Language = "UTF-8";
            StringBuilder lines = new StringBuilder();
            StringBuilder Result = new StringBuilder();
            DataTable dt = SendYMDController.GetAllMForm(date, FAB);
            lines.AppendLine(line);
            if (dt.Rows.Count != 0)
            {
                line += "<br/>";
                Result.AppendLine(line);
                //*
                foreach (DataRow dr in dt.Rows)
                {
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";
                    if (dr["IsFinishedTime"].ToString() != "")
                    {
                        line += @"=""" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @""",";
                    }
                    else
                    {
                        line += @"="""",";
                    }

                    line += @"=""" + dr["UserName"].ToString() + @""",";
                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";
                    lines.AppendLine(line);
                    line += "<br/>";
                    Result.AppendLine(line);
                }
                bool okData = SendYMDController.Download_CSV(FilePath, url, lines, Language);
            }


            return line;
        }

        /// <summary>
        /// 月排成表單明細
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string Y_List_CSV(string dateString, string FAB)
        {
            string Today = "", TodayF = "", TodayL = "";
            // 獲取當前日期
            DateTime date = SendYMDController.DateString(dateString);
            // 這個月的第一天
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            // 這個月的最後一天
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            SqlCommand cmd = new SqlCommand();
            Today = "%" + date.ToString("yyyyMMdd");
            TodayF = date.ToString("yyyyMMdd") + " 00:00:00";
            TodayL = date.ToString("yyyyMMdd") + " 23:59:59";
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = Today;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = TodayF;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = TodayL;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;

            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""項目說明"",=""規格表單編號""";
            string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + date.ToString("yyyyMMdd") + "list.csv";
            string url = ConfigurationManager.AppSettings["TxtUrl"];
            string Language = "UTF-8";
            DataTable dt = SendYMDController.GetAllMForm_List(date, FAB);
            StringBuilder lines = new StringBuilder();
            StringBuilder Result = new StringBuilder();
            lines.AppendLine(line);
            if (dt.Rows.Count != 0)
            {

                line += "<br/>";
                Result.AppendLine(line);
                //*
                foreach (DataRow dr in dt.Rows)
                {
                    string ItemTypeTxt = "正常／異常類別";
                    string ItemNameTxt = dr["ItemName"].ToString();
                    string IsError = "N";
                    string ErrorMsg = "";
                    if (dr["ItemType"].ToString() == "1")
                    {
                        if (dr["ItemValue"].ToString() != "OK")
                        {
                            IsError = "Y";
                            ErrorMsg = dr["ItemValue"].ToString();
                        }
                    }
                    if (dr["ItemType"].ToString() == "4")
                    {
                        if (dr["ItemValue"].ToString() != "N/A")
                        {
                            IsError = "Y";
                            ErrorMsg = dr["ItemValue"].ToString();
                        }
                    }
                    if (dr["ItemType"].ToString() == "2")
                    {
                        ItemTypeTxt = "數字類別";
                        ItemNameTxt = ItemNameTxt + "(" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + ")";
                        try
                        {
                            float Min = float.Parse(dr["ItemMin"].ToString());
                            float Max = float.Parse(dr["ItemMax"].ToString());
                            float Value = float.Parse(dr["ItemValue"].ToString());
                            if (Value < Min)
                            {
                                IsError = "Y";
                                ErrorMsg = dr["ItemValue"].ToString() + "低於" + dr["ItemMin"].ToString() + " (" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + ")";
                            }
                            else if (Value > Max)
                            {
                                IsError = "Y";
                                ErrorMsg = dr["ItemValue"].ToString() + "高於" + dr["ItemMax"].ToString() + " (" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + ")";
                            }
                        }
                        catch
                        {
                            IsError = "Y";
                            ErrorMsg = "資料異常";
                        }
                    }
                    if (dr["ItemType"].ToString() == "3") ItemTypeTxt = "文字類別";

                    //"廠別,巡檢單號,巡檢表單名稱,巡檢時間,人員工號,巡檢人員,項目名稱,巡檢紀錄,是否異常,異常原因";
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";

                    if (dr["IsFinishedTime"].ToString() != "")
                    {
                        line += @"=""" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @""",";
                    }
                    else
                    {
                        line += @"="""",";
                    }

                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["UserName"].ToString() + @""",";
                    line += @"=""" + ItemNameTxt + @""",";//項目名稱
                                                          // line += @"=""" + dr["SerialNumber"].ToString() + @""",";
                    /*/
                    if (dr["ItemType"].ToString() == "2")//標準值
                        line += @"=""" + dr["ItemMin"].ToString() + "~" + dr["ItemMax"].ToString() + @""",";
                    else
                        line += @"=""" + "正常" + @""",";//正常
                    //*/
                    if (dr["ItemValue"].ToString() == "OK")//巡檢紀錄
                        line += @"=""" + "正常" + @""",";
                    else
                        line += @"=""" + dr["ItemValue"].ToString() + @""",";

                    //line += @"""" + ItemTypeTxt + @""",";
                    line += @"=""" + IsError + @""",";//是否異常
                    line += @"=""" + ErrorMsg + @""",";//異常原因
                    line += @"=""" + dr["ItemContent"].ToString() + @""",";//項目說明
                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";//規格表單編號

                    lines.AppendLine(line);
                    line += "<br/>";
                    Result.AppendLine(line);
                }
                bool okData = SendYMDController.Download_CSV(FilePath, url, lines, Language);
            }

            return line;
        }

        #endregion

    }
}