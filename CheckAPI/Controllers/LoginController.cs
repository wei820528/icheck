using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace CheckAPI.Controllers
{
    public class LoginController : Controller
    {
        #region Session
        //更新session
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
        //查詢session
        private string SessionSelect(string uid)
        {
            SqlCommand cmd = new SqlCommand();
            return "";
        }
        #endregion
        // GET: Login
        #region Login
        public ActionResult LoginIndex(string t)
        {
            if (t == "f")
            {
                ViewData["msg"] = "您的帳號或密碼有誤！";
                ViewData["ErrorMsg"] = ExString;
            }
            if (t == "error")
            {
                ViewData["msg"] = "不支援低於850長度頁面！";
                //  ViewData["ErrorMsg"] = ExString;
            }
            return View();
        }

        public ActionResult ValidLogin(string uid, string pwd)
        {
            
            if (uid == "" || uid == null || pwd == "" || pwd == null) return RedirectToAction("Loginindex", "Login", new { t = "f" });
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select UserID,IsAdmin,UserName,IsInfo,IsForm from users 
                where UserID=@uid 
                and UserPwd COLLATE Latin1_General_CS_AI=@pwd
                and (IsAdmin='1' or IsAdmin='2')
                ";
            cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = uid;
            cmd.Parameters.Add("@pwd", SqlDbType.VarChar).Value = pwd;
            DataTable dt = GetSQLDataTable(cmd);
            string UserID = "";
            string IsAdmin = "";
            string IsInfo = "";
            string Name = "";
            string IsForm = "";
            // string IsInfoToken = "";
            foreach (DataRow dr in dt.Rows)
            {
                UserID = dr["UserID"].ToString();
                IsAdmin = dr["IsAdmin"].ToString();
                IsInfo = dr["IsInfo"].ToString();
                IsForm = dr["IsForm"].ToString();
                break;
            }

            if (UserID != "" && (IsAdmin == "1" || IsAdmin == "2")) {
                string[] array = new string[] { UserID, pwd, IsInfo };
                string Sessioupdate = SessionUpdate(array);
                Session["LoginAccount"] = UserID;
                Session["LoginName"] = Name;
                Session["IsInfo"] = IsInfo;
                Session["IsForm"] = IsForm;
            }
            else
                return RedirectToAction("LoginIndex", new { t = "f" });
            if (IsAdmin == "2")
                Session["AdminAccount"] = UserID;
            if (IsAdmin == "1")
                Session["CommonAccount"] = UserID;
            return RedirectToAction("AdminIndex", "Admin");
        }
        private void LoginTime(string uid)
        {
            DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update users
                set LastLoginTime=@dtime
                where UserID=@uid 
                ";
            cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = uid;
            cmd.Parameters.Add("@dtime", SqlDbType.DateTime).Value = dtime;
            GetSQLNonQuery(cmd);
        }
        public ActionResult Logout()
        {
            Session.Remove("LoginAccount");
            Session.Remove("AdminAccount");
            Session.Remove("CommonAccount");
            return RedirectToAction("LoginIndex", "Login");
        }
        public string LoginTest()
        {
            try
            {
                //string Test = Session["LoginAccount"].ToString();
                return "ok";
            }
            catch
            {
                return "Error";
            }
        }
        public ActionResult DownloadApk()
        {
            FileInfo fl = new FileInfo(Url.Content("D:/Web/Check/APK/iCheck_APP_20190417.apk"));
            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = fl.Name,
                Inline = false,
            };
            Response.AppendHeader("Content-Disposition", cd.ToString());
            Response.BufferOutput = false;
            var readStream = new FileStream(fl.FullName, FileMode.Open, FileAccess.Read);
            string contentType = MimeMapping.GetMimeMapping(fl.FullName);
            return File(readStream, contentType);
        }
        #endregion

        #region Other
        private static string ExString;
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
                    ExString = "Sql OK ";
                }
                catch (Exception ex)
                {
                    ExString = ex.ToString();
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
                    ExString = "Sql OK ";
                }
            }
            catch (Exception ex)
            {
                ExString = ex.ToString();
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
                    ExString = "Sql OK ";
                }
            }
            catch (Exception ex)
            {
                ExString = ex.ToString();
            }
            return dt;
        }
        #endregion
    }
}