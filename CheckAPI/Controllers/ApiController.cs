using CheckAPI.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web.Mvc;

namespace CheckAPI.Controllers
{
    public class ApiController : Controller
    {
        #region Login
        [HttpPost]
        public string CheckLogin()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            //解包
            string cmd = "", userid = "", userpwd = "", token = "";
            dynamic array = JsonConvert.DeserializeObject(json);
            foreach (var item in array)
            {
                cmd = item.cmd;
                userid = item.userid;
                userpwd = item.userpwd;
                break;
            }
            //查Cmd
            if (cmd != "Login")
                return JsonMessage(cmd, -5, "帳號或密碼有誤!!!");
            //查使用人數
            int Count = LoginCountTest();
            if (Count > 30)
                return JsonMessage(cmd, -5, "使用人數超過30！");
            //查帳密
            DataTable dt = LoginTest(userid, userpwd);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -5, "帳號或密碼有誤！");
            string APPIsAdmin = "", UserName = "";
            foreach(DataRow dr in dt.Rows){
                APPIsAdmin = dr["APPIsAdmin"].ToString();
                UserName = dr["UserName"].ToString();
            }
            if (APPIsAdmin == "0" || APPIsAdmin == "")
                return JsonMessage(cmd, -5, "帳號或密碼有誤！！");
            //取Token
            token = GetSetSqlToken(userid);
            if (token == "")
                return JsonMessage(cmd, -3, "資料庫回傳Token失敗");
            //寫入登入時間
            LoginTime(userid);
            //回傳Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
                {"APPIsAdmin", APPIsAdmin},
                {"UserName", UserName},
                {"token", token },
                {"userid" , userid.ToUpper() }
            };
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private int LoginCountTest()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select count(UserID) from Users
                where TokenTime>SYSDATETIME()
            ");
            int count = int.Parse(GetSQLScalar(cmd, Sqlconn));
            return count;
        }
        private DataTable LoginTest(string uid, string pwd)
        {
            DataTable dt = new DataTable();
            if (pwd == null || pwd == "") return dt;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select APPIsAdmin,UserName
                from users 
                where userid=@uid
                and UserPwd COLLATE Latin1_General_CS_AI=@pwd
            ");
            cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = uid;
            cmd.Parameters.Add("@pwd", SqlDbType.VarChar).Value = pwd;
            dt = GetSQLDataTable(cmd, Sqlconn);
            //if (APPIsAdmin == "") return "帳號或密碼有誤";
            return dt;
        }        
        private string GetSetSqlToken(string uid)
        {
            string token = GetRndToken();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                update users set token=@token,
                TokenTime=@TokenTime
                where userid=@uid
                ");
            cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = uid;
            cmd.Parameters.Add("@token", SqlDbType.VarChar).Value = token;
            cmd.Parameters.Add("@TokenTime", SqlDbType.VarChar).Value = DateTime.Now.AddHours(2).ToString("yyyy-MM-dd HH:mm:ss");
            string issave = GetSQLNonQuery(cmd, Sqlconn);
            if (issave != "ok") token = "";
            return token;
        }
        private string GetRndToken()
        {
            Guid g = Guid.NewGuid();
            var token = Convert.ToBase64String(g.ToByteArray()).Replace("=", "").Replace("+", "").Replace("/", "");
            return token;
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
            GetSQLNonQuery(cmd, Sqlconn);
        }
        #endregion

        #region Logout
        [HttpPost]
        public string Logout()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            //解包
            string cmd = "", userid = "", token = "";
            dynamic array = JsonConvert.DeserializeObject(json);
            foreach (var item in array)
            {
                cmd = item.cmd;              
                token = item.token;
                break;
            }
            userid = GetSqlUserIDbyToken(token);
            if(userid != "")
            {
                SetLogoutTime(userid);
            }
            //回傳Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" }
            };
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private string SetLogoutTime(string uid)
        {
            
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                update users set TokenTime=@TokenTime
                where userid=@uid
                ");
            cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = uid;           
            cmd.Parameters.Add("@TokenTime", SqlDbType.VarChar).Value = DateTime.Now.AddHours(0).ToString("yyyy-MM-dd HH:mm:ss");
            string Result = GetSQLNonQuery(cmd, Sqlconn);
            return Result;
        }
        #endregion

        #region QueryUser
        [HttpPost]
        public string QueryUser()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                break;
            }
            //檢查命令
            if (cmd != "QueryUser")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = GetSqlqueryUser(UserID);
            //if (dt.Rows.Count == 0) return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
            };
            JArray jar = new JArray();
            string UserName = GetUserNameByID(UserID);
            jar.Add("無");
            foreach (DataRow dr in dt.Rows)
            {
                if (UserID == dr["Agent_UserID"].ToString()) continue;
                if (dr["Agent_UserID"].ToString().Length == 4)
                    if (IsFABCheck(dr["Agent_UserID"].ToString()))
                        continue;
                if (dr["Agent_UserID"].ToString() != "" && dr["Agent_UserName"].ToString() != "")
                    jar.Add(dr["Agent_UserID"].ToString() + "," + dr["Agent_UserName"].ToString());
            }
            drRoot.Add("UserID", jar);
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private DataTable GetSqlqueryUser(string UserID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select a.UserID,a.Agent_UserID,b.UserName as Agent_UserName
                from AccountAgent as a
                left join Users as b
                on a.Agent_UserID=b.UserID
                where a.UserID=@UserID
                order by a.UserID
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select a.UserID,b.UserID as Agent_UserID
                    ,b.UserName as Agent_UserName
                from AccountAgent as a
                left join Users as b
                on a.Agent_UserID=b.FAB
                where a.UserID=@UserID
                order by a.UserID
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            DataTable dt2 = GetSQLDataTable(cmd, Sqlconn);
            foreach (DataRow dr in dt2.Rows)
            {
                if (dr["Agent_UserID"].ToString() == "" && dt2.Rows.Count == 1)
                {
                    break;
                }
                else if (dr["UserID"].ToString() == UserID && dt2.Rows.Count == 1)
                {
                    break;
                }
                else
                {
                    dt.Merge(dt2);
                    break;
                }
            }              
            return dt;
        }
        private bool IsFABCheck(string TestID)
        {
            bool IsFAB = false;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select FAB
                from Factories
                where FAB=@TestID
                ";
            cmd.Parameters.Add("@TestID", SqlDbType.VarChar).Value = TestID;
            string TestWord = GetSQLScalar(cmd, Sqlconn);
            if (TestWord != "") IsFAB = true;
            return IsFAB;
        }
        private string GetUserNameByID(string UserID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select UserName
                from Users
                where UserID=@UserID
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            string UserName = GetSQLScalar(cmd, Sqlconn);
            return UserName;
        }
        #endregion

        #region QueryFAB
        [HttpPost]
        public string QueryFAB()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", UserID = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                UserID = item.UserID;
                string[] UserData;
                UserData = UserID.Split(',');
                UserID = UserData[0];
                break;
            }
            //檢查命令
            if (cmd != "QueryFAB")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string TestID = GetSqlUserIDbyToken(token);
            if (TestID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = GetSqlqueryFAB(UserID);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
            };
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                jar.Add(dr["FAB"].ToString());
            }
            drRoot.Add("FAB", jar);
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private DataTable GetSqlqueryFAB(string UserID)
        {
            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct b.FAB
                from AccountUseTable as a
                left join Tables as b
                on a.TableID = b.TableID
                where a.UserID=@UserID
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region QueryDoc
        [HttpPost]
        public string QueryDoc()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", FAB = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                FAB = item.FAB;
                break;
            }
            //檢查命令
            if (cmd != "QueryDoc")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = GetSqlqueryDoc(FAB, UserID);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
            };
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                jar.Add(dr["Doc"].ToString());
            }
            drRoot.Add("Doc", jar);
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private DataTable GetSqlqueryDoc(string FAB, string UserID)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select b.doc
                from AccountUseTable as a
                left join Datas as b
                on a.TableID = b.TableID
				left join Tables as c
				on a.TableID=c.TableID
                left join AccountUseTable as d
				on a.TableID = d.TableID
                where a.UserID=@UserID 
                and b.FAB=@FAB
                and b.AliveTime > @dtime
                and c.TableEnable='1'                
                and b.IsFinished='0'
                order by d.TableSort
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region QueryDocSort
        [HttpPost]
        public string QueryDocSort()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", FAB = "", UserID = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                FAB = item.FAB;
                UserID = item.UserID;
                string[] UserData;
                UserData = UserID.Split(',');
                UserID = UserData[0];
                break;
            }
            //檢查命令
            if (cmd != "QueryDocSort")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string TestID = GetSqlUserIDbyToken(token);
            if (TestID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = GetSqlqueryDocSort(FAB, UserID);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
            };
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                string item = dr["TableSort"].ToString() + "." + dr["TableName"].ToString() + "\n" + dr["Doc"].ToString();
                jar.Add(item);
            }
            drRoot.Add("Doc", jar);
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private DataTable GetSqlqueryDocSort(string FAB, string UserID)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct b.Doc,a.TableSort,c.TableName
                from AccountUseTable as a
                left join Datas as b
                on a.TableID = b.TableID
				left join Tables as c
				on a.TableID=c.TableID
                left join AccountUseTable as d
				on a.TableID = d.TableID
                where a.UserID=@UserID 
                and b.FAB=@FAB
                and b.AliveTime > @dtime
                and c.TableEnable='1'                
                and b.IsFinished='0'
                order by a.TableSort,b.Doc
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region QueryDataItem
        [HttpPost]
        public string QueryDataItem()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", FAB = "", Doc = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                FAB = item.FAB;
                Doc = item.Doc;
                break;
            }
            //檢查命令
            if (cmd != "QueryDataItem")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = GetSqlqueryDataItem(FAB, Doc);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            string TableName = "";
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                TableName = dr["TableName"].ToString() + @" " + dr["Doc"].ToString();
                string ItemContent = dr["ItemContent"].ToString();
                if (dr["ItemType"].ToString() == "2")
                {
                    ItemContent = ItemContent +
                        string.Format(@"({0}~{1})",
                        dr["ItemMin"].ToString(), dr["ItemMax"].ToString());
                }
                JObject jobj = new JObject
                {
                    {"Doc",dr["Doc"].ToString() },
                    {"FAB",dr["FAB"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"ItemID",dr["ItemID"].ToString() },
                    {"ItemSort",dr["ItemSort"].ToString() },
                    {"ItemName",dr["ItemName"].ToString() },
                    {"ItemContent", ItemContent},
                    {"ItemType",dr["ItemType"].ToString() },
                    {"ItemMin",dr["ItemMin"].ToString() },
                    {"ItemMax",dr["ItemMax"].ToString() }
                };
                jar.Add(jobj);
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"Doc","" },
                    {"FAB","" },
                    {"TableID","" },
                    {"ItemID","" },
                    {"ItemSort","" },
                    {"ItemName","沒有資料" },
                    {"ItemContent","" },
                    {"ItemType","" },
                    {"ItemMin","" },
                    {"ItemMax","" }
                };
                jar.Add(jobj);
            }
            jo.Add("TableName", TableName);
            jo.Add("total", dt.Rows.Count);
            jo.Add("rows", jar);
            //SaveLog(jo.ToString());
            return JsonConvert.SerializeObject(jo).ToString();
        }
        private DataTable GetSqlqueryDataItem(string FAB, string Doc)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select a.Doc,b.FAB,b.TableID,b.TableName
                    ,c.ItemID,c.ItemSort,c.ItemName,c.ItemContent,c.ItemType
                    ,c.ItemMin,c.ItemMax
                from Datas as a
				join Tables as b
				on a.TableID=b.TableID
				join TablesItem as c
				on b.TableID = c.TableID
                where b.FAB=@FAB 
                and a.Doc=@Doc
                order by c.ItemSort
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region QueryDataItemByTag
        [HttpPost]
        public string QueryDataItemByTag()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", Tag_Code = "", UserID = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                Tag_Code = item.tid;
                UserID = item.UserID;
                string[] UserData;
                UserData = UserID.Split(',');
                UserID = UserData[0];
                break;
            }
            //檢查命令
            if (cmd != "QueryDataItemByTag")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string TestID = GetSqlUserIDbyToken(token);
            if (TestID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //找看看是否已寫入
            string IsFinished = DocIsFinished(UserID, Tag_Code);
            if (IsFinished == "True")
                return JsonMessage(cmd, -1, "今天表單已完成！");
            //取FAB和Doc
            DataTable FABandDoc = GetSqlqueryDataItemByTagCode(UserID,Tag_Code);
            string FAB = "";
            string Doc = "";
            foreach (DataRow dr in FABandDoc.Rows)
            {
                FAB = dr["FAB"].ToString();
                Doc = dr["Doc"].ToString();
                break;
            }
       //     SaveLog("QueryDataItemByTag: Doc="+Doc);
            if (Doc == "")
                return JsonMessage(cmd, -1, "沒有權限或沒有這張表單！");
            //取資料
            DataTable dt = GetSqlqueryDataItem(FAB, Doc);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            string TableName = "";
            string TableDoc = "";
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                TableName = dr["TableName"].ToString();
                TableDoc = dr["Doc"].ToString();
                string ItemContent = dr["ItemContent"].ToString();
                if (dr["ItemType"].ToString() == "2")
                {
                    ItemContent = ItemContent +
                        string.Format(@"({0}~{1})",
                        dr["ItemMin"].ToString(), dr["ItemMax"].ToString());
                }
                JObject jobj = new JObject
                {
                    {"Doc",dr["Doc"].ToString() },
                    {"FAB",dr["FAB"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"ItemID",dr["ItemID"].ToString() },
                    {"ItemSort",dr["ItemSort"].ToString() },
                    {"ItemName",dr["ItemName"].ToString() },
                    {"ItemContent", ItemContent},
                    {"ItemType",dr["ItemType"].ToString() },
                    {"ItemMin",dr["ItemMin"].ToString() },
                    {"ItemMax",dr["ItemMax"].ToString() }
                };
                jar.Add(jobj);
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"Doc","" },
                    {"FAB","" },
                    {"TableID","" },
                    {"ItemID","" },
                    {"ItemSort","" },
                    {"ItemName","沒有資料" },
                    {"ItemContent","" },
                    {"ItemType","" },
                    {"ItemMin","" },
                    {"ItemMax","" }
                };
                jar.Add(jobj);
            }
            jo.Add("error", "0");
            jo.Add("TableName", TableName);
            jo.Add("TableDoc", TableDoc);
            jo.Add("total", dt.Rows.Count);
            jo.Add("rows", jar);
            //SaveLog(jo.ToString());
            return JsonConvert.SerializeObject(jo).ToString();
        }
        private DataTable GetSqlqueryDataItemByTagCode(string UserID,string Tag_Code)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select b.FAB,b.Doc
                from AccountUseTable as a
                left join Datas as b
                on a.TableID=b.TableID
                left join TagCodeUseTable as c
                on b.TableID = c.TableID
                where b.IsFinished='0'
                and b.AliveTime > SYSDATETIME()
                and a.UserID=@UserID
                and c.Tag_Code=@Tag_Code
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
      
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        private string DocIsFinished(string UserID, string Tag_Code)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select b.IsFinished
                from AccountUseTable as a
                left join Datas as b
                on a.TableID=b.TableID
                left join TagCodeUseTable as c
                on b.TableID = c.TableID
                where b.IsFinished='0'
                and b.AliveTime > SYSDATETIME()
                and a.UserID=@UserID
                and c.Tag_Code=@Tag_Code
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            string IsFinished = GetSQLScalar(cmd, Sqlconn);
            return IsFinished;
        }
        #endregion

        #region UpdateDataItem
        [HttpPost]
        public string UpdateDataItem()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", Agent_UserID = "", rows = "";
            try
            {
                foreach (var item in array)
                {
                    cmd = item.cmd.ToString();
                    token = item.token.ToString();
                    Agent_UserID = item.Agent_UserID.ToString();
                    rows = item.rows.ToString();
                    break;
                }
            }
            catch(Exception ex)
            {
                return JsonMessage(cmd, -1, ex.ToString());
            }
            //檢查命令
            if (cmd != "UpdateDataItem")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //*檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //*/
            string reason = "", NextDoc = "N";
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            ItemData itemData = new ItemData();
            rows = "" + rows + "";
            dynamic RowDatas = JsonConvert.DeserializeObject(rows);
            SqlConnection conn = new SqlConnection(Sqlconn);
            conn.Open();
            SqlTransaction tran = conn.BeginTransaction();
            bool CommitFlag = true;
            string Doc = "";
            foreach (var Row in RowDatas)
            {
                itemData.FAB = Row.FAB;
                itemData.Doc = Row.Doc;
                Doc = Row.Doc;
                itemData.ItemID = Row.ItemID;
                itemData.ItemValue = Row.ItemValue;
                itemData.Create_at = dtime;               
                reason = GetSqlupdateDataItem(itemData, tran, conn);
                if (reason != "1") {
                    //reason = "資料項目存取錯誤";
                    CommitFlag = false; break;
                }
            }
            if (CommitFlag && Doc != "")
            {
                //
                reason = SqlInsertData(Doc, dtime, UserID, Agent_UserID, tran, conn);
                if (reason != "1") {
                    //reason = "資料存取錯誤";
                    CommitFlag = false;
                }
            }
            if (CommitFlag)
                tran.Commit();
            else
                tran.Rollback();                
            conn.Dispose();
            int error = 0;
            if (reason == "1")
            {
                reason = "ok";
                NextDoc = GetNextDoc(Doc);
            }
            else {
                error = -1;
            }
               
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", error },
                {"reason", reason },
                {"NextDoc", NextDoc }
            };
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        //
        private string GetSqlupdateDataItem(ItemData itemData, SqlTransaction tran, SqlConnection conn)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.Transaction = tran;
            cmd.CommandText = @"
                select AliveTime,IsFinished 
                from Datas 
                where Doc=@Doc 
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = itemData.Doc;
            SqlDataReader sdr = cmd.ExecuteReader();
            cmd.Dispose();
            //DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            DataTable dt = new DataTable();
            dt.Load(sdr);
            DateTime dtime = DateTime.Now;
            DateTime AliveTime = DateTime.Now.AddHours(-1);
            bool IsFinished = false;
            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                }
                catch
                {
                    AliveTime = DateTime.Now.AddHours(-1);
                }
                try
                {
                    IsFinished = bool.Parse(dr["IsFinished"].ToString());
                }
                catch
                {
                    IsFinished = false;
                }               
                break;
            }
            //檢查時間是否OK 
            if (dtime > AliveTime) return "時間已過期";
            //檢查是否存過檔了
            if (IsFinished) return "已經存過檔了";
            //存入DataItem
            cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.Transaction = tran;
            cmd.CommandText = @"
                insert into DatasItem 
                ( Doc, ItemID, ItemValue, Create_at) 
                values 
                (@Doc,@ItemID,@ItemValue,@Create_at) 
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = itemData.Doc;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = itemData.ItemID;
            cmd.Parameters.Add("@ItemValue", SqlDbType.VarChar).Value = itemData.ItemValue;
            cmd.Parameters.Add("@Create_at", SqlDbType.VarChar).Value = dtime.ToString("yyyy-MM-dd HH:mm:ss");
            //string Result = GetSQLNonQuery(cmd, Sqlconn);
            string Result = "";
            try
            {
                Result = cmd.ExecuteNonQuery().ToString();
               
            }
            catch (Exception ex)
            {
                Result = ex.ToString();
                SaveLog(Result);
            }
            cmd.Dispose();
            return Result;
        }
        //
        private string SqlInsertData(string Doc, string dtime, string UserID, string Agent_UserID, SqlTransaction tran, SqlConnection conn)
        {
            //存入DataItem
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.Transaction = tran;
            cmd.CommandText = @"
                update Datas 
				set IsFinished=1,IsFinishedTime=@dtime,UserID=@UserID,Agent_UserID=@Agent_UserID
                where Doc=@Doc 
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Agent_UserID", SqlDbType.VarChar).Value = Agent_UserID;
            string Result = "";
            try
            {
                Result = cmd.ExecuteNonQuery().ToString();
            }
            catch (Exception ex)
            {
                Result = ex.ToString();
                SaveLog(Result);
            }
            cmd.Dispose();
            //判斷FAB是否滿
            return Result;
        }
        private string GetNextDoc(string Doc)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select b.Tag_Code
                from Datas as a
                left join TagCodeUseTable as b
                on a.TableID = b.TableID
                where a.Doc = @Doc
                ");
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            string Tag_Code = GetSQLScalar(cmd, Sqlconn);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select Doc
                from TagCodeUseTable as a
                left join Datas as b
                on a.TableID = b.TableID
                where a.Tag_Code=@Tag_Code
                and b.IsFinished = '0'
                and b.AliveTime > GETDATE()
                ");
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            string NextDoc = GetSQLScalar(cmd, Sqlconn);
            if (NextDoc == "")
                return "N";
            else
                return NextDoc;
        }
        #endregion

        #region QueryFAB_All
        [HttpPost]
        public string QueryFAB_All()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                break;
            }
            //檢查命令
            if (cmd != "QueryFAB_All")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = QueryFAB_AllList(UserID);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
            };
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                jar.Add(dr["FAB"].ToString());
            }
            drRoot.Add("FAB", jar);
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private DataTable QueryFAB_AllList(string UserID)
        {
            DataTable dt = new DataTable();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct FAB
                from Factories
                order by FAB
                ");
            dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region QueryTable
        [HttpPost]
        public string QueryTable()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", FAB = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                FAB = item.FAB;
                break;
            }
            //檢查命令
            if (cmd != "QueryTable")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取資料
            DataTable dt = QueryTableList(FAB);
            if (dt.Rows.Count == 0)
                return JsonMessage(cmd, -3, "資料庫找不到資料！");
            //寫入Json
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", 0 },
                {"reason", "ok" },
            };
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                jar.Add(//dr["FAB"].ToString() + "," +
                     dr["TableID"].ToString() + "," +
                     dr["TableName"].ToString());
            }
            drRoot.Add("Table", jar);
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private DataTable QueryTableList(string FAB)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select FAB,TableID,TableName 
                from Tables 
                where TableEnable='1'
                and FAB=@FAB
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region UpdateTableByTag
        [HttpPost]
        public string UpdateTable()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", FAB = "", Tag_Code = ""
                , TableID = "", Comment = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                FAB = item.FAB;
                Tag_Code = item.Tag_Code;
                TableID = item.TableID;
                Comment = item.Comment;
                break;
            }
            //檢查命令
            if (cmd != "UpdateTable")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //*檢查Token
            if (GetSqlUserIDbyToken(token) == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //*/
            string reason = NewTagCodeProcess(Tag_Code, FAB, TableID, Comment);

            //Json寫入
            int error = 0;
            if (!reason.Contains("ok")) error = -1;
            return JsonMessage(cmd, error, reason);
        }
        private string NewTagCodeProcess(string Tag_Code, string FAB, string TableID, string Comment)
        {
            string result = "";
            if (Tag_Code == "" || FAB == "" || TableID == "")
            {
                return "資料不得為空值";
            }
            string DateMonth = GetDateMonth(TableID);
            SqlCommand cmd = new SqlCommand();
        /*    cmd.CommandText = @"
                Select Tag_Code from TagCodeUseTable
                where Tag_Code=@Tag_Code
                and DateMonth=@DateMonth
            "; */
            cmd.CommandText = @"
                Select TableID from TagCodeUseTable
                where TableID=@TableID And Tag_Code=@Tag_Code
                ";

            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            //   cmd.Parameters.Add("@DateMonth", SqlDbType.VarChar).Value = DateMonth;
            string TestTableID = GetSQLScalar(cmd, Sqlconn);
            if (TestTableID == "")
            {
                cmd = new SqlCommand();
                cmd.CommandText = @"
                    insert into TagCodeUseTable
                    ( Tag_Code, FAB, TableID, Comment, DateMonth)
                    values
                    (@Tag_Code,@FAB,@TableID,@Comment,@DateMonth)
                ";
                cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
                cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
                cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
                cmd.Parameters.Add("@DateMonth", SqlDbType.VarChar).Value = DateMonth;
                result = GetSQLNonQuery(cmd, Sqlconn);
                if (result == "ok")
                    result = "新增資料成功";
                else
                    result = "新增資料失敗";
            }
            else
            {
                cmd = new SqlCommand();
                cmd.CommandText = @"
                    update TagCodeUseTable
                    set FAB=@FAB,Comment=@Comment,DateMonth=@DateMonth
                    where TableID=@TableID and Tag_Code=@Tag_Code
                ";
                cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
                cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
                cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
                cmd.Parameters.Add("@DateMonth", SqlDbType.VarChar).Value = DateMonth;
                result = GetSQLNonQuery(cmd, Sqlconn);
                if (result == "ok")
                    result = "更新資料成功";
                else
                    result = "NewTagCodeProcess更新資料失敗";
            }
            return result;
        }
        private string GetDateMonth(string TableID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                Select WeeklyCycle,MonthCycle,YearCycle from Tables
                where TableID=@TableID
            ";
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            DataTable dt =GetSQLDataTable(cmd, Sqlconn);
            string DateMonth = "W";
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["WeeklyCycle"].ToString()!="0")
                    DateMonth = "W";
                else if (dr["MonthCycle"].ToString() != "0")
                    DateMonth = "D";
                else if (dr["YearCycle"].ToString() != "0")
                    DateMonth = "M";
            }
            return DateMonth;
        }
        #endregion

        #region 變更密碼
        [HttpPost]
        public string ChangePwd()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "", OldPwd = "", NewPwd = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                OldPwd = item.OldPwd;
                NewPwd = item.NewPwd;
                break;
            }
            //檢查命令
            if (cmd != "ChangePwd")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //*檢查Token
            string UserID = GetSqlUserIDbyToken(token);
            if (UserID == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //檢查密碼是否正確
           
            DataTable dt = LoginTest(UserID, OldPwd);
            string APPIsAdmin = "";
            foreach (DataRow dr in dt.Rows)
            {
                APPIsAdmin = dr["APPIsAdmin"].ToString();
                break;
            }
            if (APPIsAdmin == "")
                return JsonMessage(cmd, -5, "舊的密碼有誤！");
            //寫入新的密碼
            string reason = NewPwdProcess(UserID, NewPwd);
            //Json寫入
            int error = 0;
            if (!reason.Contains("ok")) error = -1;
            return JsonMessage(cmd, error, reason);
        }
        private string NewPwdProcess(string UserID ,string NewPwd)
        {
            if (UserID == "" || NewPwd == "") return "資料不得為空值";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                Update Users set UserPwd=@NewPwd
                where UserID=@UserID
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@NewPwd", SqlDbType.VarChar).Value = NewPwd;
            string result = GetSQLNonQuery(cmd, Sqlconn);
            if (result == "ok") result = "修改密碼成功";
            return result;
        }
        #endregion

        #region
        [HttpPost]
        public string GetMsg()
        {
            //接收
            Stream req = Request.InputStream;
            req.Seek(0, SeekOrigin.Begin);
            string json = new StreamReader(req).ReadToEnd();
            //SaveLog(json);
            json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            string cmd = "", token = "";
            foreach (var item in array)
            {
                cmd = item.cmd;
                token = item.token;
                break;
            }
            //檢查命令
            if (cmd != "GetMsg")
                return JsonMessage(cmd, -1, "錯誤命令！");
            //*檢查Token
            string FAB = GetSqlFABbyToken(token);
            if (FAB == "")
                return JsonMessage(cmd, -1, "憑證過期請重新登入！");
            //取得留言
            DataTable dt = GetMsgBySql(FAB);
            //Json寫入
            int error = 0;
            string reason = "";
            foreach (DataRow dr in dt.Rows)
            {
                reason += dr["Comment"].ToString() + "\n";
            }
            //if (reason == "") reason = "無留言";
            return JsonMessage(cmd, error, reason);
        }
        private DataTable GetMsgBySql(string FAB)
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select Comment 
                from MessageList
                where StartTime < @today
                and EndTime > @today
                and (FAB=@FAB or FAB='ALL')
                order by sn desc
                ");
            cmd.Parameters.Add("@today", SqlDbType.VarChar).Value = today;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        #endregion

        #region Tools
        private string GetSqlUserIDbyToken(string token)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select UserID,TokenTime  
                from users 
                where token COLLATE Latin1_General_CS_AI=@token
                ", token);
            cmd.Parameters.Add("@token", SqlDbType.VarChar).Value = token;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            string UserID = "";
            DateTime TokenTime = new DateTime();
            foreach (DataRow dr in dt.Rows)
            {
                UserID = dr["UserID"].ToString();
                if (UserID == "") return "";
                TokenTime = DateTime.Parse(dr["TokenTime"].ToString());
            }
            if (Convert.ToDateTime(TokenTime) < DateTime.Now)
            {
                UserID = "";
            }
            return UserID;
        }
        private string GetSqlFABbyToken(string token)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select FAB,TokenTime  
                from users 
                where token COLLATE Latin1_General_CS_AI=@token
                ", token);
            cmd.Parameters.Add("@token", SqlDbType.VarChar).Value = token;
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            string FAB = "";
            DateTime TokenTime = new DateTime();
            foreach (DataRow dr in dt.Rows)
            {
                FAB = dr["FAB"].ToString();
                if (FAB == "") return "";
                TokenTime = DateTime.Parse(dr["TokenTime"].ToString());
            }
            if (Convert.ToDateTime(TokenTime) < DateTime.Now)
            {
                FAB = "";
            }
            return FAB;
        }
        private string JsonMessage(string cmd, int error, string reason)
        {
            JObject drRoot = new JObject
            {
                {"cmd", cmd },
                {"error", error },
                {"reason", reason }
            };
            return JsonConvert.SerializeObject(drRoot).ToString();
        }
        private void SaveLog(string json)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                insert into Json_Log 
                ( sendjson, create_at) 
                values 
                (@sendjson,@create_at) 
                ");
            cmd.Parameters.Add("@sendjson", SqlDbType.VarChar).Value = json;
            cmd.Parameters.Add("@create_at", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            GetSQLNonQuery(cmd, Sqlconn);
        }
        private void ErrorLog(string json, string ex)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                insert into Error_Log 
                ( sendjson, ex, create_at) 
                values 
                (@sendjson,@ex,@create_at) 
                ");
            cmd.Parameters.Add("@sendjson", SqlDbType.VarChar).Value = json;
            cmd.Parameters.Add("@ex", SqlDbType.VarChar).Value = ex;
            cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            GetSQLNonQuery(cmd, Sqlconn);
        }
        #endregion

        #region SQLCmd
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        private string GetSQLScalar(SqlCommand cmd, string Sqlconn)
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
        private string GetSQLNonQuery(SqlCommand cmd, string Sqlconn)
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
        private DataTable GetSQLDataTable(SqlCommand cmd, string Sqlconn)
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
        #endregion
    }
}