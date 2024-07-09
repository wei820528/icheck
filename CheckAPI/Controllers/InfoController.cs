using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;
using CheckAPI.SettingAll;

namespace CheckAPI.Controllers
{
    public class InfoController : Controller
    {
        ApiController oApi = new ApiController();
       // private string RenewDate = ConfigurationManager.AppSettings["UpNewDate"]; //"版本更新日期 2023-05-18";//新增資料庫欄位儲存版本table
        private string RenewDate = MSSQL.RenewDate;
        #region 鎖帳號密碼用
        //權限人數管理
        private string[] InfoPeople()
        {
            int people = 7;
            // int SQLpeople = 1;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                  SELECT COUNT(IsInfo)
                  FROM [iCheck].[dbo].[Users]
                  where IsInfo='true'
                ";
            string TotalRow = GetSQLScalar(cmd);
            string[] array = new string[] { };
            if (people >= int.Parse(TotalRow))
            {

                array = new string[] { "ok", "符合人數" };

            }
            else
            {
                array = new string[] { "error", "不符合人數，請刪除" + (int.Parse(TotalRow) - people).ToString() + "人" };

            }
            return array;
        }

        //查詢權限 session 0 UserID, 1 IsInfo
        private string InfoSelect(string[] array)
        {
            string IsInfoToken = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select IsInfo
                from Users 
				where IsInfo=@IsInfo and UserID =@UserID
                ");

            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = array[0];
            cmd.Parameters.Add("@IsInfo", SqlDbType.VarChar).Value = array[1];
            DataTable dt = GetSQLDataTable(cmd);
            foreach (DataRow dr in dt.Rows)
            {
                IsInfoToken = dr["IsInfo"].ToString().Trim();

            }
            return IsInfoToken;
        }
        //更新認證 session
        //array[]  0 UserID, 1 pwd, 2 IsInfo
        private string SessionUpdate(string[] array)
        {

            string Result = "";
            string sessionSql = Session.SessionID.ToString();
            DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            string where = "";
            if (array[2].ToString() == "True")
            {
                where = ",IsInfoToken=@IsInfoToken,IsInfoTokenTime = @IsInfoTokenTime";
                cmd.Parameters.Add("@IsInfoToken", SqlDbType.VarChar).Value = sessionSql;
                cmd.Parameters.Add("@IsInfoTokenTime", SqlDbType.DateTime).Value = dtime;
            }
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = array[0].ToString();
            cmd.Parameters.Add("@Token", SqlDbType.VarChar).Value = sessionSql;
            cmd.Parameters.Add("@TokenTime", SqlDbType.DateTime).Value = dtime;
            cmd.CommandText = string.Format(@"
                update users set 
                Token=@Token,TokenTime=@TokenTime {0}
                where UserID=@UserID
                ", where);

            Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "更新資料成功";
            }
            else
            {
                Result = "更新資料失敗";
            }
            return Result;
        }
        //查詢session 0 UserID, 1 IsInfo
        private string SessionSelect(string[] array)
        {
            // string UserID = "";
            //  string IsInfo = "";
            string IsInfoToken = "";
            string sessionSql = Session.SessionID.ToString();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select IsInfoToken
                from Users 
				where IsInfo=@IsInfo and UserID =@UserID
                ");

            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = array[0];
            cmd.Parameters.Add("@IsInfo", SqlDbType.VarChar).Value = array[1];
            DataTable dt = GetSQLDataTable(cmd);

            foreach (DataRow dr in dt.Rows)
            {
                IsInfoToken = dr["IsInfoToken"].ToString().Trim();

            }
            return IsInfoToken;
        }
        #endregion
        #region 鎖IP使用
        //Get Ip Address
        protected string GetIPAddress()//Get別人ip
        {
            System.Web.HttpContext context = System.Web.HttpContext.Current;
            string ipAddress = context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

            if (!string.IsNullOrEmpty(ipAddress))
            {
                string[] addresses = ipAddress.Split(',');
                if (addresses.Length != 0)
                {
                    return addresses[0];
                }
            }
            return context.Request.ServerVariables["REMOTE_ADDR"];
        }
        private string GetFABbyIP(string IP_Address)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select count(SN)
                from Info_List
                ");
            string Count = GetSQLScalar(cmd);
            if (int.Parse(Count) > 3) return "";
            // if (Count != "3") return "";
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select FAB
                from Info_List
                where IP_Address=@IP_Address
                ");
            cmd.Parameters.Add("@IP_Address", SqlDbType.VarChar).Value = IP_Address;
            string FAB = GetSQLScalar(cmd);
            if (int.Parse(Count) == 0) return "";
            return FAB;
        }

        #endregion

        // GET: Info
        public string GetIsInfoToken(){
            string TotalRow = "";
            if (Request.Form["UserID"] != null && Request.Form["IsInfoToken"] != null) {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = string.Format(@"
                select count(IsInfoToken)
                from Users 
				where IsInfoToken=@IsInfoToken and UserID =@UserID
                ");

                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Request.Form["UserID"].ToString();
                cmd.Parameters.Add("@IsInfoToken", SqlDbType.VarChar).Value = Request.Form["IsInfoToken"].ToString();

                TotalRow =GetSQLScalar(cmd);
                if (int.Parse(TotalRow) != 0) {
                    return "ok";
                }

            }
            return "error";

        }
        public ActionResult Info()
        {
            ViewData["RenewDate"] = RenewDate;
            if (InfoPeople()[0] == "error")
            {
                ViewData["error"] = "請更改看板權限";
                return View();
            }
            if ((Request.QueryString["IsInfo"] == null || Request.QueryString["UserID"] == null) && (Request.Form["uid"] == null || Request.Form["pwd"] == null))
            {
                ViewData["error"] = "請輸入帳號密碼";
                return View();
            }
            string sessionSql = Session.SessionID.ToString();
            string IsInfo = "",UserID = "", pwd = "", all_1 = "",FAB = "",IsInfoToken = "";
           // IsInfoToken = "";
            ViewData["FAB"] = GetFAB(FAB);
            ViewData["Sort"] = GetSort();
          
            if (Request.QueryString["IsInfo"] != null && Request.QueryString["UserID"] != null){
                IsInfo = Request.QueryString["IsInfo"].ToString();
                UserID = Request.QueryString["UserID"].ToString();
                ViewData["UserID"] = Request.QueryString["UserID"].ToString();
            }
            if (    Request.Form["uid"] != null && Request.Form["pwd"] != null){
                UserID = Request.Form["uid"].ToString();
                pwd = Request.Form["pwd"].ToString();
                IsInfo = InfoSelect(new string[] { UserID, pwd });
                ViewData["UserID"] = Request.Form["uid"].ToString();
            }
            if (IsInfo != "True")
            {
                ViewData["error"] = "請輸入帳號密碼";

                return View();
            }
            all_1 = SessionUpdate(new string[] { UserID, pwd, IsInfo });
            ViewData["ok"] = "ok";
            ViewData["IsInfoToken"] = SessionSelect(new string[] { UserID, IsInfo });
            if (sessionSql != IsInfoToken)
            {
                ViewData["error"] = "請輸入帳號密碼";

                return View();
            }


            //  string sid2= Request.Params["sid"];
            //string IP_Address = GetIPAddress();
            //if (IP_Address=="::1")

            //       try { FAB = GetFABbyIP(IP_Address); } catch { FAB = ""; }
            //       if (FAB == "") return RedirectToAction("Info_Error", "Info");
            // ViewData["IP"] = IP_Address;

            return View();
        }       
        private string GetFAB(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct FAB
                from Factories
                order by FAB
                ");
            DataTable dt = GetSQLDataTable(cmd);
            string Result = "";
            if (FAB == "ALL")
                Result += string.Format(@"<option value='ALL' selected>ALL</option>");
            else
                Result += string.Format(@"<option value='ALL'>ALL</option>");
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["FAB"].ToString() == FAB)
                    Result += string.Format(@"<option value='{0}' selected>{0}</option>", dr["FAB"].ToString());
                else
                    Result += string.Format(@"<option value='{0}'>{0}</option>", dr["FAB"].ToString());
            }
            return Result;
        }
        private string GetSort()
        {
            return "";
        }
        //
        public ActionResult Info_Error()
        {
            string IP_Address = GetIPAddress();
            ViewData["IP"] = IP_Address;
            return View();
        }


        [HttpGet]//目前所有FAB
        public string AllFactories()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                    select FAB
                    from Factories
                    ");
            string sSQL = cmd.CommandText.ToString();
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                jar.Add(dr["FAB"].ToString().Trim());
            }
            jo.Add("FAB", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        [HttpPost]//主資料表
        public string GetItemList(string selectFAB)
        {
            string IP_Address = GetIPAddress();
        //    string FAB = "";
         //   try { FAB = GetFABbyIP(IP_Address); } catch { FAB = ""; }
         //   if (FAB == "") return JsonErrorMessage();
            //
            string FABLine = "";
            
            if (selectFAB != "ALL") FABLine = string.Format(@" and a.FAB=@FAB");
            if (selectFAB == "FAB8C" || selectFAB == "FAB8D") FABLine = @"and (a.FAB='FAB8C' or a.FAB='FAB8D')";
            SqlCommand cmd = new SqlCommand();

            cmd.CommandText = string.Format(@"
                    select a.FAB, a.Doc,b.TableName, a.IsFinishedTime, isnull(a.UserID,b.UserID) as UserID, c.UserName
                    from datas a
                    left join tables b on b.TableID=a.TableID
                    left join users c on c.UserID=isnull(a.UserID,b.UserID)
                    where a.AliveTime between @TodayF and @TodayL
                     {0}
                    ",FABLine);
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd") + " 00:00:00";
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd") + " 23:59:59";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = selectFAB;
            string sSQL = cmd.CommandText.ToString();
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            string error_flag = "0";
            int normal_count = 0;
            int error_count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                string flag_ok = "1";

                string IsFinishedTime = dr["IsFinishedTime"].ToString();
                try
                {
                    IsFinishedTime = DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                }
                catch { }
                if (IsFinishedTime.Equals(string.Empty))
                {
                    flag_ok = "0";
                    error_count++;
                } else
                {
                    flag_ok = "1";
                    normal_count++;
                }
                JObject jobj = new JObject
                {
                    {"FAB",dr["FAB"].ToString() },
                    {"doc",dr["Doc"].ToString() },
                    {"TableName",dr["TableName"].ToString() },
                    {"IsFinishedTime",IsFinishedTime },
                    {"UserName", dr["UserName"].ToString()},
                    {"UserID",dr["UserID"].ToString() },
                    {"ok_flag",flag_ok.ToString() }
                };
                jar.Add(jobj);
            }
            if (normal_count == 0 && error_count == 0) {
                jo.Add("null_count", 100);
            }
            normal_count = dt.Rows.Count - error_count;
            jo.Add("today", DateTime.Now.ToString("yyyy-MM-dd"));
            jo.Add("error_flag", error_flag);
            jo.Add("total", dt.Rows.Count + "");
            jo.Add("normal_count", normal_count);
            jo.Add("error_count", error_count);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }      
        [HttpPost]
        public string ChangeFlag(string Tag_Code, string ok_flag)
        {
            string ok_flag_change = "1";
            if (ok_flag == "1") ok_flag_change = "0";
            string dtime = DateTime.Now.ToString("yyyy-MM-dd");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                update NonNotes_List
                set ok_flag=@ok_flag
                where date=@dtime
                and Tag_Code=@Tag_Code
                ");
            cmd.Parameters.Add("@ok_flag", SqlDbType.VarChar).Value = ok_flag_change;
            cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            string Result = GetSQLNonQuery(cmd);
            return Result;
        }
        private string JsonErrorMessage()
        {
            JObject drRoot = new JObject
            {
                {"error_flag", "2" }
            };
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        #region SQLCmd
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        private string GetSQLScalar(SqlCommand cmd)
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

        private string GetSQLNonQuery(SqlCommand cmd)
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
            catch 
            {
                //string e = ex.ToString();
                Result = "";
            }
            return Result;
        }
        private DataTable GetSQLDataTable(SqlCommand cmd)
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
                string Error = ex.ToString();
            }
            return dt;
        }

        #endregion


    }
}