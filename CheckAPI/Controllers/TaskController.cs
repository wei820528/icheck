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
    public class TaskController : Controller
    {
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        #region 清除設定
        public string Factories_Complete_ClearAll(string FAB) 
        {
            string ok = "";
            if (FAB != "" && FAB != null) { } 
            else FAB = "";
            //判斷是否今年第一天
            ok+="年"+Factories_Complete_Y_Clear(FAB); //重製月CSV廠區

            //判斷是否這個月第一天
            ok+="月"+Factories_Complete_M_Clear(FAB);//重製年CSV廠區
            //重製單天
            ok += "日"+Factories_Complete_Clear(FAB);
            return ok;
        }
        //清除日
        public string Factories_Complete_Clear(string FAB)
        {

            string where = "";
            if (FAB != "" && FAB != null)  where = "WHERE FAB=@FAB";
            else FAB = "";
            DateTime today = DateTime.Now;
            string ok = "";
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = "";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Complete", SqlDbType.Bit).Value = 0;
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = today.ToString("yyyy-MM-dd");
            cmd.CommandText = string.Format(@"
                UPDATE Factories
                SET Complete_Time = @Complete_Time,Complete = @Complete,CreateDate = @CreateDate
                {0}", where);
            ok = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return ok;
        }
        //清除月
        public string Factories_Complete_M_Clear(string FAB)
        {

            string where = "";
            if (FAB != "" && FAB != null) where = "WHERE FAB=@FAB";
            else FAB = "";
            DateTime today = DateTime.Now;
            string ok = "";
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = "";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Complete", SqlDbType.Bit).Value = 0;
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = today.ToString("yyyy-MM-dd");
            cmd.CommandText = string.Format(@"
                UPDATE Factories
                SET CompleteM_Time = @Complete_Time,CompleteM = @Complete,CreateDate = @CreateDate
                {0}", where);
            ok = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return ok;
        }

        //清除年
        public string Factories_Complete_Y_Clear(string FAB)
        {

            string where = "";
            if (FAB != "" && FAB != null) where = "WHERE FAB=@FAB";
            else FAB = "";
            DateTime today = DateTime.Now;
            string ok = "";
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = "";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Complete", SqlDbType.Bit).Value = 0;
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = today.ToString("yyyy-MM-dd");
            cmd.CommandText = string.Format(@"
                UPDATE Factories
                SET CompleteY_Time = @Complete_Time,CompleteY = @Complete,CreateDate = @CreateDate
                {0}", where);
            ok = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return ok;
        }

        #endregion


        #region 排程出單
        //Task
        public string Task()
        {
            DateTime today = DateTime.Now;
            Factories_Complete_ClearAll("");//全部清除

            string Result = "";
           
            MSSQL.SaveLog(today.ToString("yyyyMMdd")+"準備產生單子");
            if (DateSetting.HolidayTest(today) != "") return "今天是假日";
            string Year = today.ToString("yyyy");
            string Month = today.ToString("MM");
            //每星期幾的例行表單建立
            string Weekly = today.DayOfWeek.ToString("d");//星期2 
            if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly + ",週未假日不出單";
            Result += "星期" + Weekly + "<br/>";
            DataTable dt = GetWeeklyList(Weekly);

            if (dt.Rows.Count == 0) Result += "星期" + Weekly + "無出單<br/>"; 
            else {
                foreach (DataRow dr in dt.Rows)
                {
                    MSSQL.SaveLog("日:"+dr["TableID"].ToString() + "," + dr["TableName"].ToString());
                    Result += dr["TableID"].ToString() + "," + dr["TableName"].ToString() + "<br/>";
                    Result += CreateData(dr["FAB"].ToString(), dr["TableID"].ToString(), dr["TableName"].ToString()) + "<br/>";
                    // TableID + "_" + DateTime.Now.ToString("yyyyMMdd");p
                }
            }
            //分隔線
            Result += "<hr />";
            //每月某日的例行表單建立
            string Daily = today.ToString("dd");//02號
            Result += Month + "月" + Daily + "號<br/>";
            dt = GetDailyList(Daily);
            if (dt.Rows.Count == 0) Result += Month + "月" + Daily + "號無出單<br/>";
            else {
                foreach (DataRow dr in dt.Rows)
                {
                    MSSQL.SaveLog("月:" + dr["TableID"].ToString() + "," + dr["TableName"].ToString());
                    Result += dr["TableID"].ToString() + "," + dr["TableName"].ToString() + "<br/>";
                    Result += CreateData(dr["FAB"].ToString(), dr["TableID"].ToString(), dr["TableName"].ToString()) + "<br/>";
                }
            }
            //分隔線
            Result += "<hr />";
            //每年某月的例行表單建立
            //string Month = today.ToString("MM");//10月
            Result += Year + "年" + Month + "月<br/>";
            dt = GetYearList(Month,"");
            if (dt.Rows.Count == 0) Result += Year + "年" + Month + "月無出單<br/>";
            else {
                foreach (DataRow dr in dt.Rows)
                {
                    MSSQL.SaveLog("年:" + dr["TableID"].ToString() + "," + dr["TableName"].ToString());
                    Result += dr["TableID"].ToString() + "," + dr["TableName"].ToString() + "<br/>";
                    Result += CreateMonthData(dr["FAB"].ToString(), dr["TableID"].ToString(), dr["TableName"].ToString()) + "<br/>";
                    //TableID + "_" + DateTime.Now.ToString("yyyyMM") + "M";
                }
            }

           
            //廠區變false，檢查當天是否有新單子，有的話歸零
            string FABUpdate = UpDateTrue();
            return Result;
        }

        private string UpDateTrue() {
            //先判斷是否小於當天，是的話全更新，
            DateTime today = DateTime.Now;
            string OK = "";
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                SELECT count([FAB_ID])
                  FROM  Factories
                  where CreateDate=@CreateDate
                ")
            };
            cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = "";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = "";
            cmd.Parameters.Add("@CompleteTrue", SqlDbType.Bit).Value = 1;
            cmd.Parameters.Add("@CompleteFalse", SqlDbType.Bit).Value = 0;
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = today.ToString("yyyy-MM-dd");
            string countFAB = MSSQL.GetSQLScalar(cmd,Sqlconn);
            if (countFAB == "0")
            {
                //全部都不是今天，全部初始化
                cmd.CommandText = string.Format(@"
                UPDATE Factories SET Complete_Time = @Complete_Time ,Complete = @CompleteFalse,CreateDate = @CreateDate");
                OK = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            }
            else {
                //今天的表單，查看是否有新增表單，有的話重置Complete_Time、Complete
                cmd.CommandText = string.Format(@"
                select a.FAB
                From iCheck.dbo.[Factories] as a  join
	                (SELECT *
		                  FROM Datas
		                  where CreateDate=@CreateDate  and IsFinished='0'
	                  )as b  on a.FAB=b.FAB
                where a.Complete='1' and b.CreateDate=@CreateDate and b.IsFinished='0'
                group by a.FABo
                ");
                DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                foreach (DataRow dr in dt.Rows)
                {
                    cmd.Parameters["@FAB"].Value = dr["FAB"].ToString();
                    cmd.CommandText = string.Format(@"
                        UPDATE Factories SET Complete_Time = @Complete_Time ,Complete = @CompleteFalse,CreateDate = @CreateDate  
                        where FAB =@FAB
                    ");
                    OK = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
                }
                //更新不屬於今天的
                cmd.CommandText = string.Format(@"
                UPDATE Factories SET Complete_Time = @Complete_Time ,Complete = @CompleteFalse,CreateDate = @CreateDate   where CreateDate !=@CreateDate");
                OK = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            }
            //
            return OK;
        }

        //每週星期幾的List
        private DataTable GetWeeklyList(string Weekly)
        {
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                select FAB,TableID,TableName 
                from Tables
                where TableEnable='1'
                and WeeklyCycle like @Weekly
                ")
            };
            cmd.Parameters.Add("@Weekly", SqlDbType.VarChar).Value = "%" + Weekly + "%";
            DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        //每月幾號的List
        private DataTable GetDailyList(string Daily)
        {
            if (Daily.Length < 2) Daily = "0" + Daily;
            SqlCommand cmd = new SqlCommand
            {
                CommandText = string.Format(@"
                select FAB,TableID,TableName 
                from Tables
                where TableEnable='1'
                and MonthCycle like @Daily
                ")
            };
            cmd.Parameters.Add("@Daily", SqlDbType.VarChar).Value = "%" + Daily + "%";
            DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        //每年幾月的List
        private DataTable GetYearList(string Month,string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            if (Month.Length < 2) Month = "0" + Month;
            cmd.Parameters.Add("@Month", SqlDbType.VarChar).Value = "%" + Month + "%";
            string FABWhere = "";
            if (FAB !="")
            {
                FABWhere = " and FAB=@FAB";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }
           
            cmd.CommandText = string.Format(@"
                select FAB,TableID,TableName 
               from Tables
                where TableEnable='1'
                and YearCycle like @Month {0}
                ",FABWhere);
           
            DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        //建表單
        private string CreateData(string FAB, string TableID, string TableName)
        { 
            //判斷是不是建過單了
            string SearchDoc = TableID + "_" + DateTime.Now.ToString("yyyyMMdd") + "%";
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
            string CreateDate = DateTime.Now.ToString("yyyy-MM-dd");
            string AliveTime = CreateDate + " 20:00:00";
            string Doc = TableID + "_" + DateTime.Now.ToString("yyyyMMdd"); //+ "_" + RndNumber;

            string Result = AddDatasOne(Doc,FAB,TableID,AliveTime,CreateDate);// GetSQLNonQuery(cmd, Sqlconn);
            if (Result == "ok") { 
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
        private string AddDatasOne(string Doc,string FAB,string TableID,string AliveTime, string CreateDate) {
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
        private string UpFactoriesOne(string FAB)
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
        private string CreateMonthData(string FAB, string TableID, string TableName)
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
            string AliveTime = DateTime.Now.AddDays(1 - DateTime.Now.Day).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd") + " 23:59:59";
            //string AliveTime = CreateDate + " 23:59:59";
            string Doc = TableID + "_" + DateTime.Now.ToString("yyyyMM") + "M";// + "_" + RndNumber;
            string Result = AddDatasOne(Doc, FAB, TableID, AliveTime, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));// GetSQLNonQuery(cmd, Sqlconn);
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
        #endregion
        #region 排程寄單
        //20230418新寫法 給所有廠區的人
        /// <summary>
        /// 新的寄信方式
        /// </summary>
        /// <returns></returns>
        public string NewSendMail() {
            DateTime ToDay = DateTime.Now;
            if (DateSetting.HolidayTest(ToDay) != "") return "今天是假日";
            string Weekly = ToDay.DayOfWeek.ToString("d");//星期2
            //return "今天是星期" + Weekly;
            if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly;
            string Result = "超過時間";
            int HH = int.Parse(DateTime.Now.ToString("HH"));
            if (HH >= int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
               Result = "超過時間";
            }
            else if (HH < int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                Result= NewMailText();
            }
            return Result;
        }

        private string  NewMailText() {
            string Result = "";
            string BC1 = "background-color:#E8FFFF";
            string BC2 = "background-color:#FAFAFA";
            DateTime ToDay = DateTime.Now;
            //搜尋廠區
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                SELECT FAB
                  FROM Factories
                where CreateDate=@CreateDate and Complete=@Complete
            "
            };
            //判斷未完成的廠區
            // cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = ToDay.ToString("yyyy-MM-dd HH:mm:ss");
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = ToDay.ToString();
            cmd.Parameters.Add("@Complete", SqlDbType.Bit).Value = 0;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = "";
            DataTable FABDb = MSSQL.GetSQLDataTable(cmd, Sqlconn);

            foreach (DataRow FABDr in FABDb.Rows)
            {
                string Body = "";
                string Body2 = "<table width='" + ConfigurationManager.AppSettings["Mail_Send_Width"] + "%' align='center' style='border: 3px #cccccc solid;padding:1px;' rules='all' cellpadding='8' >";
                Body2 += string.Format(@"<tr style='background-color: {0} ;color: {1} ;text-align:center;'><td  colspan='4'> {2} {3}</td></tr></table><br>",
                    ConfigurationManager.AppSettings["Mail_Send_BC"], ConfigurationManager.AppSettings["Mail_Send_C"], FABDr["FAB"].ToString(), ConfigurationManager.AppSettings["Mail_Send"]);
                string Subject = "" + FABDr["FAB"].ToString() + ConfigurationManager.AppSettings["Mail_Send_Subject"];
                cmd.Parameters["@FAB"].Value = FABDr["FAB"].ToString();
                //抓廠區內的所有表單
                //寫成mailBody
                cmd.CommandText = @"
                    SELECT a.Doc,a.TableID,b.TableName,b.TableType
                      FROM Datas as a left join [Tables] as b on a.TableID=b.TableID
                      where a.IsFinished=0 and a.CreateDate=@CreateDate and a.FAB=@FAB
                      order by b.TableType asc,b.TableID asc,a.Doc asc
                ";
                DataTable FABDatasDb = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                if (FABDatasDb.Rows.Count != 0)
                {
                    string color = "";
                    //主管
                    int countPeople = 0;//幾筆資料
                    int countP = 0;//行
                    int column = 0;//欄
                    int maxColspan = 12;//行最多幾個
                                        //cmd.CommandText = @"
                                        //select UserID,UserName,UserMail,FAB
                                        //from Users
                                        //where IsBoss='1' and FAB=@FAB
                                        //order by FAB
                                        //";
                    cmd.CommandText = @"
                    select UserID,UserName,UserMail,FAB
                    from Users
                    where (APPIsAdmin='3') and FAB=@FAB
                    order by FAB
                    ";
                    DataTable FABBossDb = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                    List<string> MailList = new List<string>();
                    Body += "<table align='center'  style='border: 3px #cccccc solid;padding:1px;' rules='all' cellpadding='8' >";
                    if (FABBossDb.Rows.Count!=0) {
                        int a = FABBossDb.Rows.Count;
                       
                        Body += @"<tr style='background-color: #009FCC;text-align:center;'><td  colspan='" + maxColspan + "'>廠區主管</td></tr>";
                        foreach (DataRow FABUserDr in FABBossDb.Rows)
                        {
                            countPeople++;

                            if (countP == maxColspan)
                            {
                                countP = 0;//行
                                column++;//欄+1
                                Body += "</tr>";
                            }
                            if (countP == 0)
                            {

                                if (column % 2 == 0) color = BC1;
                                else if (column % 2 == 1) color = BC2;
                                Body += "<tr style='" + color + "'>";//設定欄位
                            }

                            if (FABBossDb.Rows.Count == countPeople)
                            {
                                int countPEnd = (FABBossDb.Rows.Count % maxColspan); //取得最後一欄的行數
                                if (countPEnd == 0) { }
                                else
                                {
                                    Body += "<td>" + FABUserDr["UserName"].ToString() + "</td>";
                                    countPEnd = maxColspan - countPEnd;
                                    Body += "<td colspan='" + countPEnd + "'></td>";
                                }
                            }
                            else
                            {
                                Body += "<td>" + FABUserDr["UserName"].ToString() + "</td>";

                            }
                            countP++;//行+1
                          
                            MailList.Add(FABUserDr["UserMail"].ToString());
                        }
                        //  Body += "</table>";
                        Body += @"<tr style='background-color: #FFFFFF;text-align:center;'><td  colspan='" + maxColspan + "'></td></tr>";
                    }

                    //抓廠區內的所有成員
                    countP = 0;
                    column = 0;//欄
                    countPeople = 0;//幾筆資料
                                    //cmd.CommandText = @"
                                    //select UserID,UserName,UserMail,FAB
                                    //from Users
                                    //where (APPIsAdmin='1' or APPIsAdmin='2') and FAB=@FAB
                                    //order by FAB
                                    //";
                    cmd.CommandText = @"
                    select UserID,UserName,UserMail,FAB
                    from Users
                    where (APPIsAdmin='1' or APPIsAdmin='2') and FAB=@FAB
                    order by FAB
                    ";
                    DataTable FABUserDb = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                    //List<string> MailList = new List<string>();
                    //Body += "<table  style='border: 3px #cccccc solid;padding:5px;' rules='all' cellpadding='10' >";
                    if (FABUserDb.Rows.Count !=0) {
                        Body += @"<tr style='background-color: #FFF3DE;text-align:center;'><td  colspan='" + maxColspan + "'>相關同仁</td></tr>";

                        foreach (DataRow FABUserDr in FABUserDb.Rows)
                        {
                            countPeople++;
                            if (countP == maxColspan)
                            {
                                countP = 0;//行
                                column++;//欄+1
                                Body += "</tr>";
                            }
                            if (countP == 0)
                            {
                                if (column % 2 == 0) color = BC1;
                                else if (column % 2 == 1) color = BC2;
                                Body += "<tr style='" + color + "'>";//設定欄位
                            }

                            if (FABUserDb.Rows.Count == countPeople)
                            {
                                int countPEnd = (FABUserDb.Rows.Count % maxColspan); //取得最後一欄的行數
                                if (countPEnd == 0) { }
                                else
                                {
                                    Body += "<td>" + FABUserDr["UserName"].ToString() + "</td>";
                                    countPEnd = maxColspan - countPEnd;
                                    Body += "<td colspan='" + countPEnd + "'></td>";
                                }

                            }
                            else
                            {
                                Body += "<td>" + FABUserDr["UserName"].ToString() + "</td>";
                            }
                            countP++;//行+1

                            MailList.Add(FABUserDr["UserMail"].ToString());
                        }
                      

                    }
                    Body += "</table><br>";
                    if (FABDatasDb.Rows.Count != 0)
                    {
                        //未完成的表單
                        Body += "<table align='center' style='border: 3px #cccccc solid;padding:1px;' rules='all' cellpadding='8'>";
                        Body += @"<tr style='background-color: #404040;color:#FAFAFA;text-align:center;'><td  colspan='4'>未完成表單</td></tr>";
                        Body += "<tr style='background-color: #FFF3DE'><td>表單類別</td><td>表單ID</td><td>表單名稱</td><td>表單全名</td></tr>";
                        int CountB = 0;//顏色

                        foreach (DataRow FABDatasDr in FABDatasDb.Rows)
                        {
                            if (CountB == 0)
                            {
                                CountB = 1;
                                color = BC1;
                            }
                            else
                            {
                                CountB = 0;
                                color = BC2;
                            }
                            Body += "<tr style='" + color + "'>";
                            Body += "<td>" + FABDatasDr["TableType"].ToString() + "</td>";
                            Body += "<td>" + FABDatasDr["TableID"].ToString() + "</td>";
                            Body += "<td>" + FABDatasDr["TableName"].ToString() + "</td>";
                            Body += "<td>" + FABDatasDr["Doc"].ToString() + "</td>";
                            Body += "</tr>";
                        }
                        Body += "</table><br>";



                        string MailOk = MailController.SendMailByGmail(MailList, Subject, Body2+ Body);
                        //Result += Body + "寄給" + BossData["UserName"].ToString() + MailOk + "<br/><br/>";
                        MailList.Clear();
                        Result += Body2 + "<div style='text-align: center;'>" + Subject + "<br>未完成表當如下:<br></div>" + Body;

                       
                    }
                    Result += "<HR style='FILTER: alpha(opacity = 0, finishopacity = 100, style = 1)' width='80%' color=#987cb9 SIZE=3><br><br><br><br>";
                }

            }
            return Result;
        }

        public string SendMailTask()
        {
            string Result = "", Body = "",Subject = "";
            DateTime today = DateTime.Now;
            //if (HolidayTest() != "") return "今天是假日";
            if (DateSetting.HolidayTest(today) != "") return "今天是假日";
            string Weekly = today.DayOfWeek.ToString("d");//星期2
            //return "今天是星期" + Weekly;
            if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly;
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                select UserID,UserName,UserMail,FAB
                from Users
                where (APPIsAdmin='1' or APPIsAdmin='2')
                order by FAB
            "
            };
            DataTable UserTable = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            foreach (DataRow UserData in UserTable.Rows)
            {
                cmd = new SqlCommand
                {
                    CommandText = @"
                    select a.Doc,a.FAB,c.TableName
                    from Datas as a
                    left join AccountUseTable as b
                    on a.TableID = b.TableID
                    left join Tables as c
                    on a.TableID = c.TableID
                    where a.IsFinished='0'
                    and a.AliveTime > SYSDATETIME()
                    and b.UserID=@UserID
                "
                };
                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserData["UserID"].ToString();
                DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);               
                 
                Subject = "有未完成的巡檢表單";
                foreach (DataRow dr in dt.Rows)
                {

                    Body = dr["TableName"].ToString() + " 未點檢<br/>";
                    if (dt.Rows.Count != 0)
                    {
                        cmd = new SqlCommand
                        {
                            CommandText = @"
                            SELECT UserName,UserMail,FAB
                            FROM [iCheck].[dbo].[Users]
                            where (APPIsAdmin='1' or APPIsAdmin='2')
                            and FAB= @UserFAB
                            "
                        };
                        cmd.Parameters.Add("@UserFAB", SqlDbType.VarChar).Value = UserData["FAB"].ToString();
                        DataTable dFAB = MSSQL.GetSQLDataTable(cmd, Sqlconn);

                        List<string> MailList = new List<string>();
                        foreach (DataRow UserMail in dFAB.Rows)
                        {
                            MailList.Add(UserMail["UserMail"].ToString());
                        }
                        string MailOk = MailController.SendMailByGmail(MailList, Subject, Body);
                        Result += "寄給" + UserData["UserName"].ToString() + MailOk + "<br/>" + Body;
                        MailList.Clear();
                    }
                }

            }
            if (Result == "") Result = "各部門成員均已完成，沒寄出通知信</br>";

            Result += "------------------------------------------------------------------------------------------</br></br>發給各門主管通知信</br></br>";
            cmd = new SqlCommand
            {
                CommandText = @"
                select UserID,UserName,UserMail,FAB
                from Users
                where (IsBoss ='true')
                order by FAB
            "
            };
            DataTable BossTable = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            int boss = 0;
            foreach (DataRow BossData in BossTable.Rows) {
                boss++;
                Subject = boss+"."+BossData["UserName"].ToString() + "的區域有未完成的巡檢表單<br/>";
                Body = boss + "." + BossData["UserName"].ToString() + "的區域有未完成的巡檢表單<br/>";
                cmd = new SqlCommand
                {
                    CommandText = @"
                        select a.Doc,a.FAB,c.TableName
                        from Datas as a
                        left join AccountUseTable as b
                        on a.TableID = b.TableID
                        left join Tables as c
                        on a.TableID = c.TableID
                        where a.IsFinished='0'
                        and a.AliveTime > SYSDATETIME()
                        and c.FAB=@FAB
                    "
                };
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = BossData["FAB"].ToString();
                DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                int i = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    i++;
                    Body += "&nbsp;&nbsp;&nbsp;" + i+"."+dr["TableName"].ToString() + " 未點檢<br/>";
                   
                }

                if (dt.Rows.Count != 0)
                {
                    DataTable dFAB = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                    List<string> MailList = new List<string>
                    {
                        BossData["UserMail"].ToString()
                    };
                    string MailOk = MailController.SendMailByGmail(MailList, Subject, Body);
                    Result += Body + "寄給" + BossData["UserName"].ToString() + MailOk + "<br/><br/>";
                    MailList.Clear();
                }
            }
            if (Result == "") Result += "該部門不需要通知信件</br>";
            return Result;
        }
        #endregion

        #region 簽核用資料
        /// <summary>
        /// 手動用
        /// </summary>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public string SendTxt(string TxtDate)
        {
            string Result = "";
            DateTime today = SendYMDController.DateString(TxtDate); //DateTime.Now;
            bool OKholiday = DateSetting.FinalDay(TxtDate);
            //string ok = "";
            if (OKholiday)
            {
                Result = "今天是這個月最後一天平日，";
                //執行月CSV產生
                Result += TaskYMDController.C_M_CSV_ALL(TxtDate) + "<br/>";
            }
            else{
                Result += "今天不是這個月最後一天平日，";
            }
           
            // SendSystemMCSV(TxtDate,"All");
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                select FAB
                from Factories"
            };
            DataTable dt =  MSSQL.GetSQLDataTable(cmd, Sqlconn);

            foreach (DataRow dr in dt.Rows)
            {
                Result += dr["FAB"].ToString() + "<br/>";
                Result += TaskOld.SendTxt2(dr["FAB"].ToString(), TxtDate) + "<br/>";
                TaskOld.SendTxt1(dr["FAB"].ToString(), TxtDate);
            }
            return Result;
        }
        /// <summary>
        /// 每月整合
        /// </summary>
        /// <param name="TxtDate">日期</param>
        /// <param name="DataDate">All/DD/MM</param>
        /// <returns></returns>
        public string SendSystemMCSV(string TxtDate,string DataDate) {
            DateTime dateTime = DateTime.Now;
            string dateTimeFormat = "yyyy-MM-dd HH:mm:ss";
            // 嘗試解析日期時間字串
            if (DateTime.TryParseExact(TxtDate, dateTimeFormat, null, System.Globalization.DateTimeStyles.None, out DateTime parsedDateTime))
            {
                dateTime = DateTime.Parse(TxtDate);
            }
            else {
                TxtDate = "";
            }
            string Result = "無需要產出的廠區";
            string yy = dateTime.ToString("yyyy");
            string MM = dateTime.ToString("MM");
            string dd = dateTime.ToString("dd");
            SqlCommand cmd = new SqlCommand();
            try {

                //取得廠區
                cmd.CommandText = @"select FAB from Factories";
                DataTable FABAll = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                if (FABAll.Rows.Count != 0)
                {
                    foreach (DataRow FABOne in FABAll.Rows)
                    {
                        //FAB,TableID,TableName 
                        //去找出這個月有哪些單符合
                        DataTable M = GetYearList(MM, FABOne["FAB"].ToString());
                        //開始製作單
                        if (M.Rows.Count > 0)
                        {
                            Result = "";
                            //like Table ID
                            string TableIDIN = "";
                            string TableNameIN = "";
                            foreach (DataRow dr in M.Rows)
                            {
                                //廠區裡面的月
                                TableIDIN += "'" + dr["TableID"].ToString() + "',";
                                TableNameIN += "'" + dr["TableName"].ToString() + "',";
                                //FAB,TableID,TableName 
                            }
                            if (TableIDIN != "") TableIDIN = TableIDIN.Substring(0, TableIDIN.Length - 1);
                            if (TableNameIN != "") TableNameIN = TableNameIN.Substring(0, TableNameIN.Length - 1);
                             Result += TaskOld.SendMAll(FABOne["FAB"].ToString(), TxtDate, DataDate, TableIDIN, TableNameIN);
                        }
                    }

                }              
                return Result;
            }
            catch  {
                return Result;
            }

        }


        /// <summary>
        /// 自動用
        /// </summary>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public string SendSystem(string TxtDate)
        {
            // DateTime dateTime = DateTime.Now;
            //日期判斷
            DateTime dateTime = SendYMDController.DateString(TxtDate); //DateTime.Now;
            string textDay = dateTime.ToString("yyyy-MM-dd");
            string Weekly = dateTime.DayOfWeek.ToString("d");
            string M_CSV_Text = DateTime.DaysInMonth(dateTime.Year, dateTime.Month).ToString();// ConfigurationManager.AppSettings["M_CSV_Text"];
            string dd = dateTime.ToString("dd");
            string Result = "";
            bool HolidayBoolData = MSSQL.HolidayBool(TxtDate);
            SqlCommand cmd = new SqlCommand();
            DataTable dt = new DataTable();
            Result = dateTime.ToString("yyyyMMdd") + "<br>";
            Result += "排成狀態<br>";
            if (!HolidayBoolData)
            {
                cmd.CommandText = @"select FAB from Factories where Complete='0'";
                dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                if (dt.Rows.Count != 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        // Result += dr["FAB"].ToString() + "廠區狀態<br>";
                        Result = TextCsv.All_CSV_All(dr["FAB"].ToString(), TxtDate, Result,"all");
                    }
                   
                }
                else Result += "無需要產出的廠區";


             


                //Result = dateTime.ToString("yyyyMMdd")+"<br>";
                //Result += "日排成狀態<br>";
                //cmd.CommandText = @"select FAB from Factories where Complete='0'";
                //dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                //if (dt.Rows.Count != 0)
                //{

                    
                //    //執行日CSV產生
                //    foreach (DataRow dr in dt.Rows)
                //    {
                //        // Result += dr["FAB"].ToString() + "廠區狀態<br>";
                //        Result = TextCsv.All_CSV_D(dr["FAB"].ToString(), TxtDate, Result);
                //    }
                //}
                //else Result += "無需要產出的廠區";
                ////月部分
                //cmd.CommandText = @" select FAB from Factories  where CompleteM='0'";
                //dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                //bool OKholiday = DateSetting.FinalDay(textDay);
                //Result += "月排成狀態<br>";
                //if (OKholiday)
                //{
                //    Result += "今天是這個月最後一天平日<br>";
                //    if (dt.Rows.Count != 0)
                //    {
                //        //執行月CSV產生
                //        foreach (DataRow dr in dt.Rows)
                //        {
                //            // Result += dr["FAB"].ToString() + "廠區狀態<br>";
                //            Result = TextCsv.All_CSV_M(dr["FAB"].ToString(), TxtDate, Result);
                //        }
                //    }
                //    else  Result += "無需要產出的廠區";
                //}
                //else Result += "今天不是這個月最後一天";
            }
            return Result;
        }
        //更新
        private string Factories_Complete(string FAB)
        {
            DateTime today = DateTime.Now;
            string ok = "";
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = today.ToString("yyyy-MM-dd HH:mm:ss");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Complete", SqlDbType.Bit).Value = 1;
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = today.ToString("yyyy-MM-dd");
            cmd.CommandText = string.Format(@"
                UPDATE Factories
                SET Complete_Time = @Complete_Time,Complete = @Complete,CreateDate = @CreateDate
                WHERE FAB=@FAB");
            ok = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return ok;
        }

        //確認是否完成
        private string CheckComplete(string FAB, string TxtDate) {
            string ReturnOK = "";
            SqlCommand cmd = new SqlCommand();
            //查詢廠區是否完成單子
            //廠區
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            //今天時間
            DateTime today = DateTime.Now;
            if (TxtDate !="") {
                today = SendYMDController.DateString(TxtDate); //DateTime.Now;
            }
            string textDay = today.ToString("yyyy-MM-dd");
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = textDay;
            cmd.Parameters.Add("@IsFinishedFalse", SqlDbType.Bit).Value = 0;//未完成單子
            cmd.Parameters.Add("@IsFinishedTrue", SqlDbType.Bit).Value = 1;//已完成單子
            string ok = "";
            //TxtDate="" 為當天時間
            cmd.CommandText = string.Format(@"
                SELECT count([FAB])
                  FROM  Datas
                  where CreateDate=@CreateDate and FAB=@FAB and IsFinished=@IsFinishedFalse
                ");
            string CoountIsFinished = MSSQL.GetSQLScalar(cmd, Sqlconn);
            if (CoountIsFinished == "0")
            {
                //此場區全部都做了
                ReturnOK = CheckFAB(FAB, TxtDate);
                //ReturnOK = SendTxt2(FAB, TxtDate) + "<br/>";
                ok = TaskOld.SendTxt1(FAB, TxtDate);
                ReturnOK += "<br>"+ok;
                if (ok != "")
                {
                    ReturnOK+= TaskOld.SendTxt2(FAB, TxtDate);
                    Factories_Complete(FAB);
                }
            }
            else {
                ReturnOK += CheckFAB(FAB, TxtDate);
            }
            return ReturnOK;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        private string CheckFAB(string FAB, string TxtDate)
        {
            //DateTime today = DateTime.Now;
            DateTime today = SendYMDController.DateString(TxtDate);
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@IsFinishedFalse", SqlDbType.VarChar).Value = '0';
            cmd.Parameters.Add("@IsFinishedTrue", SqlDbType.VarChar).Value = '1';
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 00:00:00";
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 23:59:59";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            //未完單數量
            cmd.CommandText = @"
                select count(doc) from datas
                where doc like '%'+@Today and IsFinished = @IsFinishedFalse
                and AliveTime > @TodayF  and AliveTime < @TodayL and FAB=@FAB
                ";
            int FinishNotCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            //完單數量
            cmd.CommandText = @"
                select count(doc) from datas
                where doc like '%'+@Today and IsFinished = @IsFinishedTrue
                and AliveTime > @TodayF  and AliveTime < @TodayL and FAB=@FAB
                ";
            int FinishCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            //全部總和
            cmd.CommandText = @"select count(doc)  from datas where doc like '%'+@Today and FAB=@FAB";
            int TotalCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            string ReturnData = "<table align='center' style='border: 3px #cccccc solid;padding:1px;' rules='all' cellpadding='8' >";
            ReturnData += @"<tr style='background-color: #404040;color:#FAFAFA;text-align:center;'><td  colspan='4'>"+FAB+"</td></tr>";

            ReturnData += @"<tr style='background-color: #FFF3DE'><td  colspan='2'>應完成件數</td><td>已完成件數</td><td>尚未全部完成</td></tr>";
            ReturnData += "<tr style='background-color:#E8FFFF'>";
            ReturnData += "<td  colspan='2'>" + TotalCount + "</td>";
            ReturnData += "<td>" + FinishCount + "</td>";
            ReturnData += "<td>" + FinishNotCount + "</td>";
            ReturnData += "</tr>";
            //@"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
            cmd.CommandText = @"
                SELECT a.*,b.TableName,b.TableType,c.SerialNumber
                  FROM datas as a 
                  left join Tables as b on a.TableID=b.TableID and a.FAB=b.FAB 
                  left join TablesType as c on b.TableType=c.TableType  
                where a.doc like '%'+@Today
                and a.IsFinished = @IsFinishedFalse 
                and a.AliveTime > @TodayF
                and a.AliveTime < @TodayL and a.FAB=@FAB
                ";
            DataTable FinishNot = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            if (FinishNot.Rows.Count != 0)
            {
                ReturnData += @"<tr style='background-color: #404040;color:#FAFAFA;text-align:center;'><td  colspan='4'>" + FAB + "未完成表單</td></tr>";
                ReturnData += @"<tr style='background-color: #FFF3DE'><td>巡檢單號</td><td>巡檢表單名稱</td><td>巡檢表單類別</td><td>規格表單編號</td></tr>";
                foreach (DataRow dr in FinishNot.Rows)
                {
                    ReturnData += "<tr style='background-color:#E8FFFF'>";
                    ReturnData += "<td>" + dr["Doc"].ToString() + "</td>";
                    ReturnData += "<td>" + dr["TableName"].ToString() + "</td>";
                    ReturnData += "<td>" + dr["TableType"].ToString() + "</td>";
                    ReturnData += "<td>" + dr["SerialNumber"].ToString() + "</td>";
                    ReturnData += "</tr>";
                }
            }
            ReturnData += "</table><br><HR style='FILTER: alpha(opacity = 0, finishopacity = 100, style = 1)' width='80%' color=#987cb9 SIZE=3>";
            return ReturnData;


        }
        #endregion


        #region 20240528產生單
        //Task
        public string TaskD(string TxtDate)
        {
            //DateTime today = DateTime.Now;
            DateTime today = SendYMDController.DateString(TxtDate);
            //DateTime today = DateTime.Now;
            Factories_Complete_ClearAll("");//全部清除

            string Result = "";

            MSSQL.SaveLog(today.ToString("yyyyMMdd") + "準備產生單子");
            if (DateSetting.HolidayTest(today) != "") return "今天是假日";
            string Year = today.ToString("yyyy");
            string Month = today.ToString("MM");
            //每星期幾的例行表單建立
            string Weekly = today.DayOfWeek.ToString("d");//星期2 
            if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly + ",週未假日不出單";
            Result += "星期" + Weekly + "<br/>";
            DataTable dt = GetWeeklyList(Weekly);


            if (dt.Rows.Count == 0) Result += "星期" + Weekly + "無出單<br/>";
            else
            {
                foreach (DataRow dr in dt.Rows)
                {
                    MSSQL.SaveLog("日:" + dr["TableID"].ToString() + "," + dr["TableName"].ToString());
                    Result += dr["TableID"].ToString() + "," + dr["TableName"].ToString() + "<br/>";
                    Result += TaskAdd.CreateData(dr["FAB"].ToString(), dr["TableID"].ToString(), dr["TableName"].ToString(), today) + "<br/>";
                    // TableID + "_" + DateTime.Now.ToString("yyyyMMdd");p
                }
            }
            //分隔線
            Result += "<hr />";
            //每月某日的例行表單建立
            string Daily = today.ToString("dd");//02號
            Result += Month + "月" + Daily + "號<br/>";
            dt = GetDailyList(Daily);
            if (dt.Rows.Count == 0) Result += Month + "月" + Daily + "號無出單<br/>";
            else
            {
                foreach (DataRow dr in dt.Rows)
                {
                    MSSQL.SaveLog("月:" + dr["TableID"].ToString() + "," + dr["TableName"].ToString());
                    Result += dr["TableID"].ToString() + "," + dr["TableName"].ToString() + "<br/>";
                    Result += TaskAdd.CreateData(dr["FAB"].ToString(), dr["TableID"].ToString(), dr["TableName"].ToString(), today) + "<br/>";
                }
            }
            //分隔線
            Result += "<hr />";
            //每年某月的例行表單建立
            //string Month = today.ToString("MM");//10月
            Result += Year + "年" + Month + "月<br/>";
            dt = GetYearList(Month, "");
            if (dt.Rows.Count == 0) Result += Year + "年" + Month + "月無出單<br/>";
            else
            {
                foreach (DataRow dr in dt.Rows)
                {
                    MSSQL.SaveLog("年:" + dr["TableID"].ToString() + "," + dr["TableName"].ToString());
                    Result += dr["TableID"].ToString() + "," + dr["TableName"].ToString() + "<br/>";
                    Result += TaskAdd.CreateMonthData(dr["FAB"].ToString(), dr["TableID"].ToString(), dr["TableName"].ToString(), today) + "<br/>";
                    //TableID + "_" + DateTime.Now.ToString("yyyyMM") + "M";
                }
            }


            //廠區變false，檢查當天是否有新單子，有的話歸零
            string FABUpdate = UpDateTrue();
            return Result;
        }
        #endregion

    }
}