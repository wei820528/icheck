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
    public class MSSQL
    {
        #region SQLCmd
        public static readonly string RenewDate = "版本更新日期 20240202";
        public static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        public static string GetSQLScalar(SqlCommand cmd, string Sqlconn)
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
        public static string GetSQLNonQuery(SqlCommand cmd, string Sqlconn)
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
        public static DataTable GetSQLDataTable(SqlCommand cmd, string Sqlconn)
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
        public static void SaveLog(string json)
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
        /// <summary>
        /// 判斷是假日，true為是假日，false為不是假日
        /// </summary>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static bool HolidayBool(string TxtDate)
        {
            DateTime dateTime = SendYMDController.DateString(TxtDate); 
            string Weekly = dateTime.DayOfWeek.ToString("d");
            if (DateSetting.HolidayTest(dateTime) != "") return true;
            if (Weekly == "6" || Weekly == "0") return true;
            return false;
        }
        #endregion

    }
}
