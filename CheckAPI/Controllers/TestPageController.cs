using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CheckAPI.SettingAll;

namespace CheckAPI.Controllers
{
    public class TestPageController : Controller
    {

        #region 廠區
        public string GetFABbs()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct FAB
                    from Factories 
                    where FAB !='' 
                    order by FAB
                ");
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value='{0}' selected>{0}</option>
                        ", dr["FAB"].ToString());
                    i++;
                }
                else
                {
                    result += string.Format(@"
                        <option value='{0}'>{0}</option>
                        ", dr["FAB"].ToString());
                }
            }
            return result;
        }
        #endregion
        #region 點檢表單 測試頁面
        public ActionResult TestPage()
        {
            return View();
        }
        public string GetDocbs(string FAB, string TxtDate)
        {
            DateTime today = SendYMDController.DateString(TxtDate);
            string TodayF = today.ToString("yyyyMMdd") + " 00:00:00";
            string TodayL = today.ToString("yyyyMMdd") + " 23:59:59";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select Doc,TableID,AliveTime,IsFinished
                from Datas
                where FAB=@FAB
                and AliveTime > @TodayF
                and AliveTime < @TodayL 
                and IsFinished='0'
                order by Doc
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = TodayF;
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = TodayL;
            //cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            //cmd.Parameters.Add("@dtime2", SqlDbType.VarChar).Value = dtime;
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value='{0},{1}' selected>{0}</option>
                        ", dr["Doc"].ToString(), dr["TableID"].ToString());
                    i++;
                }
                else
                {
                    result += string.Format(@"
                        <option value='{0},{1}'>{0}</option>
                        ", dr["Doc"].ToString(), dr["TableID"].ToString());
                }
            }
            if (dt.Rows.Count == 0)
                result += @"<option value='' selected>無資料</option>";
            return result;
        }
        public string GetTestData(string FAB, string Doc, string TableID, string TxtDate)
        {
            //string s = Session["LoginAccount"].ToString();
            //string c = Session["AdminAccount"].ToString();
            //string d = Session["LoginName"].ToString();
            //     string f = Session["CommonAccount"].ToString();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select *
                from TablesItem
                where TableID=@TableID
                order by ItemSort
                ");
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            //使用Json回傳
            string result = @"";
            int i = 1;
            string ItemList = "";
            foreach (DataRow dr in dt.Rows)
            {
                //資料用
                if (dr["ItemSort"].ToString() != i.ToString())
                {
                    //改ItemSort
                    ChangeItemSortGo(TableID, dr["ItemID"].ToString(), i);
                }
                ItemList += dr["ItemID"].ToString() + ",";
                //畫面使用
                string SelectLine = "";
                if (dr["ItemType"].ToString() == "1")
                {
                    SelectLine = string.Format(@"
                        <input type=""text"" class=""form-control"" id=""{0}"" disabled/>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','OK','OK')"">正常</button>
                        </div>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','Error','')"">異常</button>
                        </div>
                        ", dr["ItemID"].ToString());
                }
                else if (dr["ItemType"].ToString() == "2")
                {
                    SelectLine = string.Format(@"
                        <input type=""number"" class=""form-control"" id=""{0}""/>
                        ", dr["ItemID"].ToString());
                }
                else if (dr["ItemType"].ToString() == "4")
                {
                    SelectLine = string.Format(@"
                        <input type=""text"" class=""form-control"" id=""{0}"" disabled/>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','OK','OK')"">正常</button>
                        </div>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','OK','N/A')"">無此裝置</button>
                        </div>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','Error','')"">異常</button>
                        </div>
                        ", dr["ItemID"].ToString());
                }
                else
                {
                    SelectLine = string.Format(@"
                        <input type=""text"" class=""form-control"" id=""{0}""/>
                        ", dr["ItemID"].ToString());
                }


                string Content = dr["ItemContent"].ToString();
                if (Content.Length > 15)
                {
                    Content = Content.Substring(0, 13) + "<br/>" + Content.Substring(13, Content.Length - 13);
                }
                result += string.Format(@"
                    <div class=""input-group col-12 col-sm-12 col-md-12 col-lg-12"" style=""margin-top:20px;"">
                        <div class=""input-group-prepend"">
                            <label class=""input-group-text"" for=""Doc"" style=""width:300px;"">
                                {1}<br/>
                                {2}
                            </label>
                        </div>
                        {3}
                    </div>
                    <input type='text' id='{0}_time' value='' style='display:none;'/>
                    ", dr["ItemID"].ToString(), dr["ItemName"].ToString(), Content, SelectLine);
                i++;
            }
            if (dt.Rows.Count > 1) ItemList = ItemList.Substring(0, ItemList.Length - 1);
            result += string.Format(@"
                <input type='text' id='ItemList' value='{0}' style='display:none;'/>
                <div class=""form-row col-12 col-sm-12 col-md-12 col-lg-12 text-center"" style=""margin-top:20px;margin-bottom:20px;"">
                    <div class=""col-12"">
                        <button type=""button"" class=""btn btn-info col-5 col-sm-5 col-md-5 col-lg-5"" onclick=""UpdateTable()"">上傳</button>
                        <button type=""button"" class=""btn btn-danger col-5 col-sm-5 col-md-5 col-lg-5"" onclick=""CloseTable()"">取消</button>
                    </div>
                </div>
                ", ItemList);
            return result;
        }
        public string UpdateDatas(string json,string TxtDate)
        {
            DateTime today = SendYMDController.DateString(TxtDate);
            string Result = "";

            string Doc = "", ItemID = "", ItemValue = "", UserID = "";
            //json = "[" + json + "]";
            dynamic array = JsonConvert.DeserializeObject(json);
            //dynamic array = json;
            foreach (var item in array)
            {
                Doc = item.Doc;
                ItemID = item.ItemID;
                ItemValue = item.ItemValue;
                UserID = Session["LoginAccount"].ToString();
                Result = UpdateDatasItem(Doc, ItemID, ItemValue, UserID, TxtDate);
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select AliveTime,IsFinished
                from Datas
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            //
           // DateTime dtime = DateTime.Now;
            DateTime AliveTime = today.AddHours(-1);
            bool IsFinished = true;
            foreach (DataRow dr in dt.Rows)
            {
                AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                IsFinished = bool.Parse(dr["IsFinished"].ToString());
                break;
            }
            //檢查時間是否OK 
            if (today > AliveTime) return "單號時間已過期";
            //檢查是否存過檔了
            if (IsFinished) return "已經存過檔了";
            //存入
            cmd = new SqlCommand();
            cmd.CommandText = @"
                update Datas set
                IsFinished='1',IsFinishedTime=@dtime,UserID=@UserID
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            cmd.Parameters.Add("@dtime", SqlDbType.DateTime).Value = today;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            Result = MSSQL.GetSQLNonQuery(cmd, MSSQL.Sqlconn);
            if (Result == "ok")
                Result = "儲存單號" + Doc + "成功";
            else
                Result = "儲存單號" + Doc + "失敗";
            return Result;
        }
        public string UpdateDatasItem(string Doc, string ItemID, string ItemValue, string UserID, string TxtDate)
        {
            DateTime today = SendYMDController.DateString(TxtDate);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select AliveTime,IsFinished
                from Datas
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = MSSQL.GetSQLDataTable(cmd, MSSQL.Sqlconn);
            //
           // DateTime dtime = DateTime.Now;
            DateTime AliveTime = today.AddHours(-1);
            bool IsFinished = true;
            foreach (DataRow dr in dt.Rows)
            {
                AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                IsFinished = bool.Parse(dr["IsFinished"].ToString());
                break;
            }
            //檢查時間是否OK 
            if (today > AliveTime) return "";
            //檢查是否存過檔了
            if (IsFinished) return "";
            //存入
            cmd = new SqlCommand();
            cmd.CommandText = @"
                insert into DatasItem
                ( Doc, ItemID, ItemValue, UserID, Create_at)
                values
                (@Doc,@ItemID,@ItemValue,@UserID,@Create_at)
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = ItemID;
            cmd.Parameters.Add("@ItemValue", SqlDbType.VarChar).Value = ItemValue;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Create_at", SqlDbType.VarChar).Value = today.ToString("yyyy-MM-dd HH:mm:ss");
            string Result = MSSQL.GetSQLNonQuery(cmd, MSSQL.Sqlconn);
            //成功 OK 失敗 ""
            return Result;
        }
        private void ChangeItemSortGo(string TableID, string ItemID, int ItemSort)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update TablesItem set
                ItemSort=@ItemSort
                where TableID=@TableID
                and ItemID=@ItemID
                ";
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = ItemID;
            cmd.Parameters.Add("@ItemSort", SqlDbType.VarChar).Value = ItemSort;
            string Result = MSSQL.GetSQLNonQuery(cmd, MSSQL.Sqlconn);
        }
        #endregion
    }
}