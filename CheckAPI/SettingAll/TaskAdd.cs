using CheckAPI.Controllers;
using CheckAPI.Models;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Http.Results;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;


namespace CheckAPI.SettingAll
{
    public class TaskAdd
    {
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        public static string CreateMonthData(string FAB, string TableID, string TableName, DateTime today)
        {
            //判斷是不是建過單了
            string SearchDoc = TableID + "_" + DateTime.Now.ToString("yyyyMM") + "M%";
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                select Doc from Datas
                where Doc like @Doc
                ")
            };
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = SearchDoc;
            string TestDoc = MSSQL.GetSQLScalar(cmd, Sqlconn);
            if (TestDoc != "") return TestDoc + "已存在";
            //建立表單
            Random rnd = new Random();
            string RndNumber = rnd.Next(101, 999).ToString();
           // string AliveTime = DateTime.Now.AddDays(1 - DateTime.Now.Day).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd") + " 23:59:59";
            string AliveTime = today.AddDays(1 - today.Day).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd") + " 23:59:59";
            //string AliveTime = CreateDate + " 23:59:59";

            string Doc = TableID + "_" + DateTime.Now.ToString("yyyyMM") + "M";// + "_" + RndNumber;
           // string Result = AddDatasOne(Doc, FAB, TableID, AliveTime, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));// GetSQLNonQuery(cmd, Sqlconn);
            string Result = AddDatasOne(Doc, FAB, TableID, AliveTime, today.ToString("yyyy-MM-dd HH:mm:ss"));// GetSQLNonQuery(cmd, Sqlconn);
            if (Result == "ok")
            {
                Result = Doc + "已建立成功</br>";
                string Result2 = UpFactoriesOne(FAB);
                if (Result2 == "ok")
                {
                    Result += FAB + "廠區以重製";

                }
            }
            else Result = Doc + "建單失敗";
            return Result;
        }
        //建表單
        public static string CreateData(string FAB, string TableID, string TableName, DateTime today)
        {
            //判斷是不是建過單了
           // string SearchDoc = TableID + "_" + DateTime.Now.ToString("yyyyMMdd") + "%";
            string SearchDoc = TableID + "_" + today.ToString("yyyyMMdd") + "%";
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                select Doc from Datas
                where Doc like @Doc
                ")
            };
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = SearchDoc;
            string TestDoc = MSSQL.GetSQLScalar(cmd, Sqlconn);
            if (TestDoc != "") return TestDoc + "已存在";
            //建立表單
            Random rnd = new Random();
            string RndNumber = rnd.Next(101, 999).ToString();
           // string CreateDate = DateTime.Now.ToString("yyyy-MM-dd");
            string CreateDate = today.ToString("yyyy-MM-dd");
            string AliveTime = CreateDate + " 20:00:00";
           // string Doc = TableID + "_" + DateTime.Now.ToString("yyyyMMdd"); //+ "_" + RndNumber;
            string Doc = TableID + "_" + today.ToString("yyyyMMdd"); //+ "_" + RndNumber;
            string Result = AddDatasOne(Doc, FAB, TableID, AliveTime, CreateDate);// GetSQLNonQuery(cmd, Sqlconn);
            if (Result == "ok")
            {
                Result = Doc + "已建立成功</br>";
                string Result2 = UpFactoriesOne(FAB);
                if (Result2 == "ok")
                {
                    Result += FAB + "廠區以重製";

                }
            }
            else Result = Doc + "建單失敗</br>";
            return Result;
        }

        private static string AddDatasOne(string Doc, string FAB, string TableID, string AliveTime, string CreateDate)
        {
            string Result = "";
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                insert into Datas
                ( Doc, FAB, TableID, AliveTime, IsFinished, CreateDate)
                values
                (@Doc,@FAB,@TableID,@AliveTime,'0',@CreateDate)
                ")
            };
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@AliveTime", SqlDbType.VarChar).Value = AliveTime;
            cmd.Parameters.Add("@CreateDate", SqlDbType.VarChar).Value = CreateDate;
            Result = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return Result;
        }
        private static string UpFactoriesOne(string FAB)
        {
            string Result = "";
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                UPDATE [dbo].[Factories]
                   SET [Complete] =0
                 WHERE FAB=@FAB
                ")
            };
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            Result = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return Result;
        }
    }
}