using CheckAPI.Models;
using CheckAPI.SettingAll;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Xml;
//2021/5/27 點檢表單部分，完成後錯誤顯示
namespace CheckAPI.Controllers
{
    [UserLogin_filter]
    public class AdminController : Controller
    {
        [Description("20230414更新CSV和點檢表單、20230501更新寄信mail、20230510更新每小時出單、20230518更新mail規則")]
        private string RenewDate = MSSQL.RenewDate; //ConfigurationManager.AppSettings["UpNewDate"]; //"版本更新日期 2023-07-24";//匯入時判斷csv編碼，並且進行寫入動作

        //"上次版本更新日期 2021-07-16";//新增資料庫欄位儲存版本table
        //DataManage改寫
        // GET: Check
        public ActionResult AdminIndex()
        {
            ViewData["RenewDate"] = RenewDate;
            return View();
        }

        #region 表單管理
        public ActionResult TableManage()
        {
            return View();
        }

        public string GetDate()
        {
            JArray jar = new JArray();
            JObject jobj = new JObject
            {
                {"Value","ALL" },
                {"Text","ALL" },
                {"selected",true }
            };
            jar.Add(jobj);

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select DISTINCT  CONVERT(varchar(100), a.AliveTime, 23) as date, a.AliveTime as data
                from Datas as a
                left join Tables as t
                on a.TableID=t.TableID
                        ");
            DataTable dt = GetSQLDataTable(cmd);
            foreach (DataRow dr in dt.Rows)
            {
                jobj = new JObject
                {
                    {"Value",dr["date"].ToString() },
                    {"Text",dr["date"].ToString() },
                };
                jar.Add(jobj);
            }



            return JsonConvert.SerializeObject(jar).ToString();
            //SqlCommand cmd = new SqlCommand();
            //cmd.CommandText = string.Format(@"
            //    select distinct FAB
            //    from Factories
            //    where FAB !='' order by FAB
            //    ");
            //DataTable dt = GetSQLDataTable(cmd);
            ////使用Json回傳
            //JArray jar = new JArray();
            //JObject jobj = new JObject
            //{
            //    {"Value","ALL" },
            //    {"Text","ALL" },
            //    {"selected",true }
            //};
            //jar.Add(jobj);
            //foreach (DataRow dr in dt.Rows)
            //{
            //    jobj = new JObject
            //    {
            //        {"Value",dr["FAB"].ToString() },
            //        {"Text",dr["FAB"].ToString() },
            //    };
            //    jar.Add(jobj);
            //}
            // return JsonConvert.SerializeObject(jar).ToString();
        }

        public string GetFAB()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct FAB
                from Factories
                where FAB !='' order by FAB
                ");
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            JObject jobj = new JObject
            {
                {"Value","ALL" },
                {"Text","ALL" },
                {"selected",true }
            };
            jar.Add(jobj);
            foreach (DataRow dr in dt.Rows)
            {
                jobj = new JObject
                {
                    {"Value",dr["FAB"].ToString() },
                    {"Text",dr["FAB"].ToString() },
                };
                jar.Add(jobj);
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string GetFABbs()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct FAB
                    from Factories 
                    where FAB !='' 
                    order by FAB
                ");
            DataTable dt = GetSQLDataTable(cmd);
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
        public string GetTypebs()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct TableType
                    from TablesType
                    order by TableType
                ");
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value='{0}' selected>{0}</option>
                        ", dr["TableType"].ToString());
                    i++;
                }
                else
                {
                    result += string.Format(@"
                        <option value='{0}'>{0}</option>
                        ", dr["TableType"].ToString());
                }
            }
            return result;
        }
        public string GetUserbs(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct UserID
                from Users 
                where FAB=@FAB
                order by UserID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value='{0}' selected>{0}</option>
                        ", dr["UserID"].ToString());
                    i++;
                }
                else
                {
                    result += string.Format(@"
                        <option value='{0}'>{0}</option>
                        ", dr["UserID"].ToString());
                }
            }
            if (dt.Rows.Count == 0)
            {
                result += string.Format(@"
                    <option value=''>NoUser</option>
                    ");
            }
            return result;
        }
        public string GetTableList(string FAB, int page, int rows)
        {
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            string FABLine = "";
            if (FAB != "ALL")
                FABLine = @"where FAB=@FAB";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select * from
                    (select ROW_NUMBER() OVER(ORDER BY FAB,TableID) AS 'RowNo'
                    ,FAB,TableID,TableName,TableType
                    ,WeeklyCycle,MonthCycle,YearCycle,TableEnable,UserID 
                    from Tables
                    {0}
                    ) as t
				where RowNo between @LeftNo and @RightNo
                ", FABLine);
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            if (FAB != "ALL")
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"select COUNT(TableID) from Tables {0}", FABLine);
            if (FAB != "ALL")
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                string EnableWord = "<span style='color:red'>停用</span>";
                if (bool.Parse(dr["TableEnable"].ToString()))
                    EnableWord = "<span style='color:blue'>啟用</span>";
                JObject jobj = new JObject
                {
                    {"FAB",dr["FAB"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"TableName",dr["TableName"].ToString() },
                    {"TableType",dr["TableType"].ToString() },
                    {"WeeklyCycle",dr["WeeklyCycle"].ToString() },
                    {"MonthCycle",dr["MonthCycle"].ToString() },
                    {"YearCycle",dr["YearCycle"].ToString() },
                    {"TableEnable",dr["TableEnable"].ToString() },
                    {"EnableWord",EnableWord },
                    {"UserID",dr["UserID"].ToString() }
                };
                jar.Add(jobj);
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"FAB","" },
                    {"TableID","" },
                    {"TableName","" },
                     {"TableType","" },
                    {"WeeklyCycle","沒有資料" },
                    {"YearCycle","" },
                    {"MonthCycle","" },
                    {"UserID","" }
                };
            }
            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string CheckTableIDDuplicate(string FAB, string TableID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select TableID 
                from Tables 
                where FAB=@FAB
                and TableID=@TableID";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            string Test = GetSQLScalar(cmd, Sqlconn);
            if (Test != "")
                return "t";
            else
                return "f";
        }
        public string AddTable(Tables tables)
        {
            string result = "";
            DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                    insert into Tables
                    ( TableID, TableName, TableType, FAB, WeeklyCycle, MonthCycle, YearCycle, TableEnable, UserID)
                    values
                    (@TableID,@TableName,@TableType,@FAB,@WeeklyCycle,@MonthCycle,@YearCycle,@TableEnable,@UserID)
                    ";
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
            cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value = tables.TableName;
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = tables.TableType;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tables.FAB;
            cmd.Parameters.Add("@WeeklyCycle", SqlDbType.VarChar).Value = tables.WeeklyCycle;
            cmd.Parameters.Add("@MonthCycle", SqlDbType.VarChar).Value = tables.MonthCycle;
            cmd.Parameters.Add("@YearCycle", SqlDbType.VarChar).Value = tables.YearCycle;
            cmd.Parameters.Add("@TableEnable", SqlDbType.Bit).Value = tables.TableEnable;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = "";
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
            {
                result = "表單新增成功";
            }
            else
            {
                result = "表單新增失敗";
            }
            return result;
        }
        public string UpdateTable(Tables tables)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update Tables set 
                TableName=@TableName,TableType=@TableType
                ,WeeklyCycle=@WeeklyCycle,MonthCycle=@MonthCycle,YearCycle=@YearCycle
                ,TableEnable=@TableEnable,FAB=@FAB,UserID=@UserID
                where FAB=@FAB 
                and TableID=@TableID
                ";
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
            cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value = tables.TableName;
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = tables.TableType;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tables.FAB;
            cmd.Parameters.Add("@WeeklyCycle", SqlDbType.VarChar).Value = tables.WeeklyCycle;
            cmd.Parameters.Add("@MonthCycle", SqlDbType.VarChar).Value = tables.MonthCycle;
            cmd.Parameters.Add("@YearCycle", SqlDbType.VarChar).Value = tables.YearCycle;
            cmd.Parameters.Add("@TableEnable", SqlDbType.VarChar).Value = tables.TableEnable;
            if (tables.UserID == null)
                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = "";
            else
                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = tables.UserID;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "更新表單資料成功";
            }
            else
            {
                Result = "更新表單資料失敗";
            }
            return Result;
        }
        public string DeleteTable(Tables tables)
        {
            string Result = "無TagCode管理清除表單";
            //先搜尋這個tables是否正確
            if (tables == null || tables.FAB == "" || tables.TableID == "") {
                return Result + "[0]";
            }
            //查詢tables是否存在
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"SELECT TableID 
                FROM Tables 
                where TableID=@TableID";
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
            DataTable dt = GetSQLDataTable(cmd);
            if (dt.Rows.Count == 0)
            {
                return Result + "[1]";
            }
            //判斷這個表單是否用到
            cmd = new SqlCommand();
            cmd.CommandText = @"
                select count(*) as count
                from(	
                    SELECT TableID
                    FROM [iCheck].[dbo].[AccountUseTable]
                    where TableID = @TableID 
                    UNION ALL
                    SELECT TableID
                    FROM  [iCheck].[dbo].[Datas]
                    where TableID = @TableID
                    UNION ALL
                    SELECT TableID
                    FROM  [iCheck].[dbo].[TablesItem]
                    where TableID = @TableID
                    UNION ALL
                    SELECT TableID
                    FROM  [iCheck].[dbo].[TagCodeUseTable]
                    where TableID = @TableID
                ) as a

            ";

            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
            dt = GetSQLDataTable(cmd);
            string count = "";
            foreach (DataRow dr in dt.Rows) {
                count = dr["count"].ToString();
            }
            if (count == "0") {
                cmd = new SqlCommand();
                cmd.CommandText = @"
                delete Tables 
                where FAB=@FAB
                and TableID=@TableID";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tables.FAB;
                cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
                Result = GetSQLNonQuery(cmd);
                cmd = new SqlCommand();
                cmd.CommandText = @"
                delete TablesItem
                where FAB=@FAB
                and TableID=@TableID";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tables.FAB;
                cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
                Result += GetSQLNonQuery(cmd);
                cmd = new SqlCommand();
                cmd.CommandText = @"
                delete AccountUseTable 
                where TableID=@TableID
                ";
                cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;

                Result += GetSQLNonQuery(cmd);

                if (Result == "okokok")
                {
                    /*清空設定用
                    string TagCode = "";
                    string TableID = "";
                    Result = null;
                    if (tables.TableID != "")
                        TableID = @"where TableID=@TableID";
                    cmd = new SqlCommand();
                    cmd.CommandText = string.Format(@"
                    select * from TagCodeUseTable
                    {0}
                    ", TableID);
                    if (tables.TableID != "")cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tables.TableID;
                    dt = GetSQLDataTable(cmd);

                    if (dt.Rows.Count == 0)
                    {
                        Result = "無TagCode管理清除表單";
                        return Result;
                    }
                    else
                    {
                        Result += "dt.Rows.Count:" + dt.Rows.Count;
                        foreach (DataRow dr in dt.Rows)
                        {
                            //  Result += "\n TageCode dr " + dr["TagCode"].ToString();
                            TagCode += "'"+dr["Tag_Code"].ToString()+"',";

                        }
                        TagCode = TagCode.Substring(0, TagCode.Length - 1);
                        Result += "\n TagCode " + TagCode;
                        //
                    }
                   // string[] TagCode2 = System.Text.RegularExpressions.Regex.Split(TagCode, @",");
                    //更新表單
                    if (TagCode == "")
                    {

                        Result = "無TagCode管理清除表單";
                        return Result;
                    }
                    else {
                       string TableTag = @" WHERE Tag_Code in ("+TagCode+@")";
                      //  Result += "\n TableTag: " + TableTag;
                         cmd = new SqlCommand();
                         cmd.CommandText = string.Format(@"
                         update TagCodeUseTable 
                         set  FAB = '' , TableID = '' , DateMonth = '' 
                         {0}", TableTag);
                        //   cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = TagCode;
                        Result += GetSQLNonQuery(cmd);

                    }
                    */
                    Result += "\n tables.TableID: " + tables.TableID;
                    Result += "\n tables.FAB: " + tables.FAB;
                    // Result += "\n tables.TableName: " + tables.TableName;
                    // Result += "\n tables.TableType: " + tables.TableType;
                    // Result += "\n tables.WeeklyCycle: " + tables.WeeklyCycle;
                    // Result += "\n tables.MonthCycle: " + tables.MonthCycle;
                    // Result += "\n tables.YearCycle: " + tables.YearCycle;
                    // Result += "\n tables.TableEnable: " + tables.TableEnable;
                    // Result += "\n tables.UserID: " + tables.UserID;
                    Result += "\n 刪除表單" + tables.TableID + ":" + tables.TableName + "成功";
                    return "OK" + Result;
                }
                else
                {
                    Result = "刪除表單" + tables.TableID + ":" + tables.TableName + "失敗";
                }
                //return Result;
            }
            else
            {
                Result = "刪除表單" + tables.TableID + ":" + tables.TableName + "失敗";
            }

            return Result;
        }
        public string GetPrincipalList(string TableID, string UserID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select UserID
                from AccountUseTable
                where TableID=@TableID
                ");
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value=''>無</option>
                        ");
                    i++;
                }
                result += string.Format(@"
                    <option value='{0}'>{0}</option>
                    ", dr["UserID"].ToString());
            }
            if (dt.Rows.Count == 0)
            {
                result += string.Format(@"
                    <option value=''>無</option>
                    ");
            }
            return result;
        }
        #endregion

        #region 表單項目管理
        public ActionResult ItemManage()
        {
            return View();
        }
        public string GetItemFAB()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct FAB
                from Tables
                order by FAB
                ");
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 1)
                {
                    JObject jobj = new JObject
                    {
                        {"Value",dr["FAB"].ToString() },
                        {"Text",dr["FAB"].ToString() },
                        {"selected",true }
                    };
                    jar.Add(jobj);
                }
                else
                {
                    JObject jobj = new JObject
                    {
                        {"Value",dr["FAB"].ToString() },
                        {"Text",dr["FAB"].ToString() },
                    };
                    jar.Add(jobj);
                }
                i++;
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string GetItemTableID(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select TableID,TableName
                from Tables
                where FAB=@FAB
                order by TableID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 1)
                {
                    JObject jobj = new JObject
                    {
                        {"Value",dr["TableID"].ToString() },
                        {"Text",dr["TableName"].ToString() },
                        {"selected",true }
                    };
                    jar.Add(jobj);
                }
                else
                {
                    JObject jobj = new JObject
                    {
                        {"Value",dr["TableID"].ToString() },
                        {"Text",dr["TableName"].ToString() }
                    };
                    jar.Add(jobj);
                }
                i++;
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                    {
                        {"Value","" },
                        {"Text","NoData" },
                        {"selected",true }
                    };
                jar.Add(jobj);
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string GetItemList(string FAB, string TableID, int page, int rows)
        {
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select * from
                    (select ROW_NUMBER() OVER(ORDER BY ItemSort) AS 'RowNo'
                    ,FAB,TableID,ItemID,ItemSort,ItemName,ItemContent,ItemType,ItemMin,ItemMax
                    from TablesItem 
                    where FAB=@FAB 
                    and TableID=@TableID
                    ) as t
				where RowNo between @LeftNo and @RightNo
                ");
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            DataTable dt = GetSQLDataTable(cmd);
            cmd = new SqlCommand();
            cmd.CommandText = @"
                select COUNT(ItemID) 
                from TablesItem 
                where FAB=@FAB
                and TableID=@TableID
                ";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                //資料用
                if (dr["ItemSort"].ToString() != i.ToString())
                {
                    //改ItemSort
                    //ChangeItemSort(TableID, dr["ItemID"].ToString(), i);
                }
                string ItemType = dr["ItemType"].ToString();
                string ItemTypeWord = "";
                if (ItemType == "1") ItemTypeWord = "正確/異常 選項";
                if (ItemType == "2") ItemTypeWord = "數字類別";
                if (ItemType == "3") ItemTypeWord = "文字類別";
                if (ItemType == "4") ItemTypeWord = "無此裝置";
                JObject jobj = new JObject
                {
                    {"FAB",dr["FAB"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"ItemID",dr["ItemID"].ToString() },
                    {"ItemSort",dr["ItemSort"].ToString() },
                    {"ItemName",dr["ItemName"].ToString() },
                    {"ItemContent",dr["ItemContent"].ToString() },
                    {"ItemType",dr["ItemType"].ToString() },
                    {"ItemMin",dr["ItemMin"].ToString() },
                    {"ItemMax",dr["ItemMax"].ToString() },
                    {"ItemTypeWord",ItemTypeWord }
                };
                jar.Add(jobj);
                i++;
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"FAB","" },
                    {"TableID","" },
                    {"ItemID","" },
                    {"ItemSort","" },
                    {"ItemName","沒有資料" },
                    {"ItemContent","" },
                    {"ItemType","" },
                    {"ItemTypeWord","" }
                };
            }
            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        private void ChangeItemSort(TablesItem tablesItem)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select FAB,TableID,ItemID,ItemSort 
                from TablesItem 
                where FAB=@FAB 
                and TableID=@TableID
                order by ItemSort
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tablesItem.FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tablesItem.TableID;
            DataTable dt = GetSQLDataTable(cmd);
            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                //資料用
                if (dr["ItemSort"].ToString() != i.ToString())
                {
                    //改ItemSort
                    ChangeItemSortGo(tablesItem.TableID, dr["ItemID"].ToString(), i);
                }
                i++;
            }
        }
        public string CheckItemIDDuplicate(string FAB, string TableID, string ItemID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select ItemID 
                from TablesItem 
                where FAB=@FAB 
                and TableID=@TableID 
                and ItemID=@ItemID
                ";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = ItemID;
            string Test = GetSQLScalar(cmd, Sqlconn);
            if (Test != "")
                return "t";
            else
                return "f";
        }
        public string AddTableItem(TablesItem tablesItem)
        {
            string result = "";
            DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                    insert into TablesItem
                    ( FAB, TableID, ItemID, ItemSort, ItemName, ItemContent, ItemType, ItemMin, ItemMax, CreateTime)
                    values
                    (@FAB,@TableID,@ItemID,@ItemSort,@ItemName,@ItemContent,@ItemType,@ItemMin,@ItemMax,@CreateTime)
                    ";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tablesItem.FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tablesItem.TableID;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = tablesItem.ItemID;
            cmd.Parameters.Add("@ItemSort", SqlDbType.VarChar).Value = tablesItem.ItemSort;
            cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = tablesItem.ItemName;
            cmd.Parameters.Add("@ItemContent", SqlDbType.VarChar).Value = tablesItem.ItemContent;
            cmd.Parameters.Add("@ItemType", SqlDbType.VarChar).Value = tablesItem.ItemType;
            cmd.Parameters.Add("@ItemMin", SqlDbType.Int).Value = tablesItem.ItemMin;
            cmd.Parameters.Add("@ItemMax", SqlDbType.Int).Value = tablesItem.ItemMax;
            cmd.Parameters.Add("@CreateTime", SqlDbType.DateTime).Value = dtime;
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
            {
                ChangeItemSort(tablesItem);
                result = "項目新增成功";
            }
            else
            {
                result = "項目新增失敗";
            }
            return result;
        }
        public string UpdateTableItem(TablesItem tablesItem)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update TablesItem set 
                ItemSort=@ItemSort,ItemName=@ItemName,ItemContent=@ItemContent
                ,ItemType=@ItemType,ItemMin=@ItemMin,ItemMax=@ItemMax
                where FAB=@FAB 
                and TableID=@TableID and ItemID=@ItemID
                ";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tablesItem.FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tablesItem.TableID;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = tablesItem.ItemID;
            cmd.Parameters.Add("@ItemSort", SqlDbType.VarChar).Value = tablesItem.ItemSort;
            cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = tablesItem.ItemName;
            cmd.Parameters.Add("@ItemContent", SqlDbType.VarChar).Value = tablesItem.ItemContent;
            cmd.Parameters.Add("@ItemType", SqlDbType.VarChar).Value = tablesItem.ItemType;
            cmd.Parameters.Add("@ItemMin", SqlDbType.Int).Value = tablesItem.ItemMin;
            cmd.Parameters.Add("@ItemMax", SqlDbType.Int).Value = tablesItem.ItemMax;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                ChangeItemSort(tablesItem);
                Result = "更新資料成功";
            }
            else
            {
                Result = "更新資料失敗";
            }
            return Result;
        }
        public string DeleteTableItem(TablesItem tablesItem)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"delete TablesItem where FAB=@FAB and TableID=@TableID and ItemID=@ItemID";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = tablesItem.FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = tablesItem.TableID;
            cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = tablesItem.ItemID;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                ChangeItemSort(tablesItem);
                Result = "刪除資料" + tablesItem.TableID + "/" + tablesItem.ItemID + "成功";
            }
            else
            {
                Result = "刪除資料" + tablesItem.TableID + "/" + tablesItem.ItemID + "失敗";
            }
            return Result;
        }
        #endregion

        #region 匯入表單頁面
        public ActionResult ImportManage()
        {
            return View();
        }

        public string PostExcelDataExport(string From,string FromName) {
            //搜尋資料表所有內容
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@" SELECT * FROM {0}", From);
            DataTable dt = GetSQLDataTable(cmd);
            string line = "";
            bool csv = false;
            StringBuilder lines = new StringBuilder();
            foreach (DataColumn DC in dt.Columns)
            {
                line +=  DC.ColumnName + @";";
            }
            line = line.Substring(0, line.Length - 1);
            lines.AppendLine(line);

            if (dt.Rows.Count != 0)
            {
                csv = true;
                foreach (DataRow dr in dt.Rows)
                {
                    line = "";
                    foreach (DataColumn DC in dt.Columns)
                    {
                        
                        line += dr[DC.ColumnName].ToString() + @";";
                    }
                    line=line.Substring(0,line.Length-1);
                    lines.AppendLine(line);
                }
            }
            if (csv)
            {
                string targetDir = System.Web.HttpContext.Current.Server.MapPath("~/SaveCSV");//路徑Txt
                string path = Path.Combine(targetDir, Path.GetFileName(FromName+"_"+DateTime.Now.ToString("yyyyMMdd") + ".csv"));
                if (!Directory.Exists(targetDir))
                {
                    //新增資料夾
                    Directory.CreateDirectory(targetDir);
                }

                //*存文字
                try
                {
                    System.IO.File.WriteAllText(path, lines.ToString(), Encoding.GetEncoding("UTF-8"));
                    try
                    {
                        return "本機" + path;
                    }
                    catch {
                        return "儲存失敗";
                    }
                    
                }
                catch (Exception ex)
                {
                    SaveLog(ex.ToString());
                    return "error";
                }
                //  Factories_Complete(FAB);//已匯出
            }
            else {
                return "找尋不到此資料表";
            }
        }

        public string PostExcelDataSave(string From, string FromName)
        {
            string Result = "";//類別名稱
            if (Session["LoginAccount"].ToString() == null || Session["LoginAccount"].ToString() == "")
                return "逾時過期請重新登入";
            DataTable Items = new DataTable();
            try
            {
                var a = Request.Files[0];
                string b = a.FileName;
                HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                HttpPostedFile file = files[0];
                Result = file.FileName;
                //HttpPostedFile file = CSVPath;
                string targetDir = System.Web.HttpContext.Current.Server.MapPath("~/UpdataCSV");//路徑Txt
                string path = Path.Combine(targetDir, Path.GetFileName(file.FileName));
                if (!Directory.Exists(targetDir))
                {
                    //新增資料夾
                    Directory.CreateDirectory(targetDir);
                    //  Request.Files[i].SaveAs(FilePath);
                }
                file.SaveAs(path);
                Items = GetDataTableCSV(path, From);
                int aa = Items.Rows.Count;
                if (aa !=0) {
                    return  SaveDataTableCSV(From, Items);
                }
            }
            catch (Exception ex)
            {
                SaveLog(ex.ToString());
                return "檔案存入錯誤";
            }
            return "檔案存入有問題!請聯絡開發者";
        }
        private string SaveDataTableCSV(string From, DataTable Items) {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@" TRUNCATE TABLE  {0}", From);
            string ok = GetSQLNonQuery(cmd);
            // DataTable DB = GetSQLDataTable(cmd);
            string Name = "",No="";
            foreach (DataColumn DC in Items.Columns)
            {
                bool DCAutoIncrement = DC.AutoIncrement;
                if (!DCAutoIncrement) {
                    string TypeD = DC.DataType.ToString();
                    if (TypeD == "System.String") cmd.Parameters.Add("@" + DC.ColumnName, SqlDbType.VarChar).Value = "";
                    else if (TypeD == "System.Int32") cmd.Parameters.Add("@" + DC.ColumnName, SqlDbType.Int).Value = 0;
                    else if (TypeD == "System.Boolean") cmd.Parameters.Add("@" + DC.ColumnName, SqlDbType.Bit).Value = 0;
                    else if (TypeD == "System.DateTime") cmd.Parameters.Add("@" + DC.ColumnName, SqlDbType.DateTime).Value = DateTime.Now;
                    Name += DC.ColumnName + ",";
                    No += "@" + DC.ColumnName + ",";
                }
            }
            if (Name =="" && No=="") {
                return "匯入資料異常，無筆數";
            }
            if (Name.Split(',').Length == No.Split(',').Length)
            {
                int AllCount=0,OkCount = 0,ErrorCount=0;
                Name = Name.Substring(0, Name.Length - 1);
                No = No.Substring(0, No.Length - 1);
                foreach (DataRow dr in Items.Rows)
                {
                    foreach (DataColumn DC in Items.Columns)
                    {
                        string TypeD = DC.DataType.ToString();
                        bool DCAutoIncrement = DC.AutoIncrement;
                        if (!DCAutoIncrement)
                        {
                            if (TypeD == "System.String") cmd.Parameters["@"+DC.ColumnName].Value = dr[DC.ColumnName].ToString();
                            else if (TypeD == "System.Int32") cmd.Parameters["@" + DC.ColumnName].Value = int.Parse(dr[DC.ColumnName].ToString());
                            else if (TypeD == "System.Boolean") cmd.Parameters["@" + DC.ColumnName].Value = bool.Parse(dr[DC.ColumnName].ToString());
                            else if (TypeD == "System.DateTime") cmd.Parameters["@" + DC.ColumnName].Value = DateTime.Parse(dr[DC.ColumnName].ToString());
                        }

                    }

                    cmd.CommandText = $@" INSERT INTO {From} ( {Name} ) VALUES ( {No} )";// string.Format(, From,Name,No);
                    ok = GetSQLNonQuery(cmd);
                    if (ok == "ok") OkCount++;
                    else  ErrorCount++;
                    AllCount++;
                }
                return "全部筆數"+ AllCount + "，匯入資料成功"+ OkCount + "筆"+"，匯入失敗"+ ErrorCount + "筆";
            }
            else
            {
                return "匯入資料異常";
            }        
        }
        private DataTable GetDataTableCSV(string txtpath,string From)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@" SELECT * FROM {0}", From);
            DataTable DB = GetSQLDataTable(cmd);
            DataTable dt = new DataTable();
            var a = DB.PrimaryKey;
            foreach (DataColumn DC in DB.Columns)
            {
                bool DCAutoIncrement = DC.AutoIncrement;
                //String
                dt.Columns.Add(DC.ColumnName, DC.DataType);
                if (DCAutoIncrement) {
                    dt.Columns[DC.ColumnName].AutoIncrement = DCAutoIncrement;
                    dt.Columns[DC.ColumnName].AutoIncrementStep = 1;
                }
                //  dt.Columns.[DC.ColumnName] = "";
            }
            string line;
            string FilePath = txtpath;
            var encoding = GetEncoding(FilePath);
            int j = 0;
            using (StreamReader file = new StreamReader(FilePath, encoding))
            {
                while ((line = file.ReadLine()) != null)
                {
                    
                    string[] FileData = line.Split(';');
                    int i = 0;
                    DataRow dr = dt.NewRow();
                    foreach (DataColumn DC in DB.Columns)
                    {
                        if (FileData[i] == DC.ColumnName) {
                            break;
                        } 
                        else {
                            string TypeD = DC.DataType.ToString();
                            if (TypeD == "System.String")  dr[DC.ColumnName] = FileData[i].Trim();
                            else if (TypeD == "System.Int32")dr[DC.ColumnName] = int.Parse(FileData[i].Trim());
                            else if (TypeD == "System.Boolean") dr[DC.ColumnName] = bool.Parse(FileData[i].Trim());
                            else if (TypeD == "System.DateTime") dr[DC.ColumnName] = DateTime.Parse(FileData[i].Trim());
                            i++;
                        }
                    }
                    if (i == DB.Columns.Count)
                    {
                        dt.Rows.Add(dr);
                    }
                    j++;
                }
            }
            return dt;
        }
        public string PostExcelType() {
            string Result = "";//類別名稱
            if (Session["LoginAccount"].ToString() == null || Session["LoginAccount"].ToString() == "")
                return "逾時過期請重新登入";
            DataTable Items = new DataTable();
            try
            {
                HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                HttpPostedFile file = files[0];
                Result = file.FileName;
                //HttpPostedFile file = CSVPath;
                string targetDir = System.Web.HttpContext.Current.Server.MapPath("~/TxtType");//路徑Txt
                string path = Path.Combine(targetDir, Path.GetFileName(file.FileName));
                if (!Directory.Exists(targetDir))
                {
                    //新增資料夾
                    Directory.CreateDirectory(targetDir);
                    //  Request.Files[i].SaveAs(FilePath);
                }
                file.SaveAs(path);
                Items = GetUpdateTypeCSV(path);
            }
            catch (Exception ex)
            {
                SaveLog(ex.ToString());
                return "檔案存入錯誤";
            }
            //存SQL
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            int SaveCount = 0,Updata=0;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                  SELECT max(SN) FROM TablesType
                ");
            int Max = int.Parse(GetSQLScalar(cmd, Sqlconn));
            cmd.Parameters.Add("@SN", SqlDbType.Int).Value = Max;
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = "";
            cmd.Parameters.Add("@SerialNumber", SqlDbType.VarChar).Value = "";
            foreach (DataRow item in Items.Rows)
            {
                string TableType = item["TableType"].ToString();
                cmd.Parameters["@TableType"].Value = item["TableType"].ToString();
                cmd.Parameters["@SerialNumber"].Value = item["SerialNumber"].ToString();
                cmd.CommandText = string.Format(@"SELECT SN FROM TablesType where TableType=@TableType");
                DataTable dt = GetSQLDataTable(cmd);
                if (dt.Rows.Count == 0)
                {
                    Max=Max+1;
                    cmd.Parameters["@SN"].Value = Max;
                    //新增
                    cmd.CommandText = string.Format(@"
                            insert into TablesType  (TableType, SerialNumber)
                            values (@TableType,@SerialNumber)
                        ");
                    string Test = GetSQLNonQuery(cmd);
                    if (Test == "ok") SaveCount++;
                }
                else {
                    //更新
                    foreach (DataRow dr in dt.Rows)
                    {
                        cmd.Parameters["@SN"].Value = int.Parse(dr["SN"].ToString());
                    }
                    cmd.CommandText = string.Format(@"
                            UPDATE TablesType
                               SET SerialNumber = @SerialNumber
                             WHERE TableType=@TableType
                    ");
                    string Test = GetSQLNonQuery(cmd);
                    if (Test == "ok") Updata++;
                }
            }
            return Result + "已更新" + Updata + "筆類別，" + SaveCount + "筆類別";

        }
        private DataTable GetUpdateTypeCSV(string txtpath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("TableType");
            dt.Columns.Add("SerialNumber");
            string line;
            string FilePath = txtpath;
            var encoding = GetEncoding(FilePath);
            using (StreamReader file = new StreamReader(FilePath, encoding))
            {
                while ((line = file.ReadLine()) != null)
                {
                    string[] FileData = line.Split(',');
                    //string a = FileData[3].Split('(')[0];
                    if (FileData[0] == "" || FileData[0] == @"類別") continue;
                    if (FileData[1] == "" || FileData[1] == @"規格單號編號") continue;
                    //if (FileData[3] == "" || FileData[3].Split('(')[0] == @"項目類別") continue;
                    try
                    {
                        DataRow dr = dt.NewRow();
                        dr["TableType"] = FileData[0].Trim();
                        dr["SerialNumber"] = FileData[1].Trim();
                        dt.Rows.Add(dr);
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            return dt;
        }
        private static string TableType = "NoType";
        public string ChangeItemType(string Type)
        {
            TableType = Type;
            return TableType;
        }
        public string PostExcelData()
        {
            string Result = TableType;//類別名稱
            //Session檢查
            if (Session["LoginAccount"].ToString() == null || Session["LoginAccount"].ToString() == "")
                return "逾時過期請重新登入";
            //取CSV資料
            DataTable Items = new DataTable();
            try
            {
                HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                HttpPostedFile file = files[0];
                //HttpPostedFile file = CSVPath;
                string targetDir = System.Web.HttpContext.Current.Server.MapPath("~/Txt");//路徑Txt
                string path = Path.Combine(targetDir, Path.GetFileName(file.FileName));
                if (!Directory.Exists(targetDir))
                {
                    //新增資料夾
                    Directory.CreateDirectory(targetDir);
                    //  Request.Files[i].SaveAs(FilePath);
                }
                file.SaveAs(path);
                Items = GetUpdateCSV(path);
            }
            catch(Exception ex)
            {
                SaveLog(ex.ToString());
                return ex.ToString();
            }
            //存SQL
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            int SaveCount = 0;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select FAB,TableID from Tables where TableType=@TableType
                ");
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = TableType;
            DataTable dt = GetSQLDataTable(cmd);
            foreach (DataRow dr in dt.Rows)
            {
                cmd = new SqlCommand();
                cmd.CommandText = string.Format(@"
                    delete TablesItem where TableID=@TableID
                    ");
                cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = dr["TableID"].ToString();
                GetSQLNonQuery(cmd);
                int i = 1;
                foreach (DataRow item in Items.Rows)
                {
                    cmd = new SqlCommand();
                    cmd.CommandText = string.Format(@"
                        insert into TablesItem 
                        ( FAB, TableID, ItemID, ItemSort, ItemName, ItemContent, ItemType, CreateTime, ItemMin, ItemMax)
                        values
                        (@FAB,@TableID,@ItemID,@ItemSort,@ItemName,@ItemContent,@ItemType,@CreateTime,@ItemMin,@ItemMax)
                    ");
                    cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = dr["FAB"].ToString();
                    cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = dr["TableID"].ToString();
                    cmd.Parameters.Add("@ItemID", SqlDbType.VarChar).Value = item["ItemID"].ToString();
                    cmd.Parameters.Add("@ItemSort", SqlDbType.VarChar).Value = i.ToString();
                    cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = item["ItemName"].ToString();
                    cmd.Parameters.Add("@ItemContent", SqlDbType.VarChar).Value = item["ItemContent"].ToString();
                    cmd.Parameters.Add("@ItemType", SqlDbType.VarChar).Value = item["ItemType"].ToString();
                    cmd.Parameters.Add("@CreateTime", SqlDbType.VarChar).Value = dtime;
                    cmd.Parameters.Add("@ItemMin", SqlDbType.VarChar).Value = item["ItemMin"].ToString();
                    cmd.Parameters.Add("@ItemMax", SqlDbType.VarChar).Value = item["ItemMax"].ToString();
                    string Test = GetSQLNonQuery(cmd);
                    if (Test == "ok") SaveCount++;
                    i++;
                }
            }
            return Result + "已更新" + (SaveCount / Items.Rows.Count) + "張表單，" + SaveCount + "筆項目";
        }
        public static string ByteArrayToHexStrNew(byte[] vabytData)
        {
            return BitConverter.ToString(vabytData).Replace("-", string.Empty).ToLower();
        }





        /// <summary>
        /// 獲取文件編碼方式
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static Encoding GetEncoding(string filename)
        {
            using (System.IO.FileStream fs = new System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                System.IO.BinaryReader br = new System.IO.BinaryReader(fs);
                byte[] buffer = br.ReadBytes(2);

                if (buffer[0] >= 0xEF)
                {
                    if (buffer[0] == 0xEF && buffer[1] == 0xBB)
                    {
                        return System.Text.Encoding.UTF8;
                    }
                    else if (buffer[0] == 0xFE && buffer[1] == 0xFF)
                    {
                        return System.Text.Encoding.BigEndianUnicode;
                    }
                    else if (buffer[0] == 0xFF && buffer[1] == 0xFE)
                    {
                        return System.Text.Encoding.Unicode;
                    }
                    else
                    {
                        return System.Text.Encoding.Default;
                    }
                }
                else
                {
                    return System.Text.Encoding.Default;
                }
            }

        }

        private DataTable GetUpdateCSV(string txtpath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ItemID");
            dt.Columns.Add("ItemName");
            dt.Columns.Add("ItemContent");
            dt.Columns.Add("ItemType");
            dt.Columns.Add("ItemMin");
            dt.Columns.Add("ItemMax");
            string line;
            string FilePath = txtpath;
            var encoding = GetEncoding(FilePath);  
            using (StreamReader file = new StreamReader(FilePath, encoding))
            {
                while ((line = file.ReadLine()) != null)
                {
                    string[] FileData = line.Split(',');
                    //string a = FileData[3].Split('(')[0];
                    if (FileData[0] == "" || FileData[0] == @"項目ID") continue;
                    if (FileData[1] == "" || FileData[1] == @"項目名稱") continue;
                    if (FileData[2] == "" || FileData[2] == @"項目說明") continue;
                    //if (FileData[3] == "" || FileData[3].Split('(')[0] == @"項目類別") continue;
                    try
                    {
                        DataRow dr = dt.NewRow();
                        dr["ItemID"] = FileData[0].Trim();
                        dr["ItemName"] = FileData[1].Trim();
                        dr["ItemContent"] = FileData[2].Trim();
                        dr["ItemType"] = FileData[3].Trim();
                        if (dr["ItemType"].ToString() == "2" && FileData.Length > 5)
                        {
                            //if (FileData[4] == "" || FileData[4] == @"項目最小值") continue;
                            //if (FileData[5] == "" || FileData[5] == @"項目最大值") continue;
                            dr["ItemMin"] = FileData[4].Trim();
                            dr["ItemMax"] = FileData[5].Trim();
                        }
                        else
                        {
                            dr["ItemMin"] = 0;
                            dr["ItemMax"] = 0;
                        }
                        dt.Rows.Add(dr);
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            return dt;
        }
        #endregion

        #region 類別管理
        public string GetTypeList()
        {
            //RowData
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select SN,TableType,SerialNumber from TablesType
                    ");
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                JObject jobj = new JObject
                {
                    {"SN",dr["SN"].ToString() },
                    {"TableType",dr["TableType"].ToString() },
                    {"SerialNumber",dr["SerialNumber"].ToString() }
                };
                jar.Add(jobj);
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"SN","" },
                    {"TableType","沒有資料" },
                    {"SerialNumber","沒有資料"}
                };
                jar.Add(jobj);
            }
            jo.Add("total", dt.Rows.Count);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string GetTableType()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select TableType
                from TablesType
                where TableType !='' 
                order by TableType
                ");
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    JObject jobj = new JObject
                    {
                        {"Value",dr["TableType"].ToString() },
                        {"Text",dr["TableType"].ToString() },
                        {"selected","true" }
                    };
                    jar.Add(jobj);
                }
                else
                {
                    JObject jobj = new JObject
                    {
                        {"Value",dr["TableType"].ToString() },
                        {"Text",dr["TableType"].ToString() }
                    };
                    jar.Add(jobj);
                }
                i++;
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string AddTableType(string TableType,string SerialNumber)
        {
            string result = "";
           // DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                    insert into TablesType
                    ( TableType,SerialNumber)
                    values
                    (@TableType,@SerialNumber)
                    ";
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = TableType;
            cmd.Parameters.Add("@SerialNumber", SqlDbType.VarChar).Value = SerialNumber;
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
            {
                result = "類別新增成功";
            }
            else
            {
                result = "類別新增失敗";
            }
            return result;
        }
        public string UpdateTableType(int SN, string TableType,string SerialNumber)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@SN", SqlDbType.Int).Value = SN;
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = TableType;
            cmd.Parameters.Add("@SerialNumber", SqlDbType.VarChar).Value = SerialNumber;
            //得到舊的TableType
            cmd.CommandText = @"
             SELECT [TableType]
             FROM [iCheck].[dbo].[TablesType]
             where SN=@SN";
            string OldTableType=GetSQLScalar(cmd, Sqlconn);
            cmd.Parameters.Add("@TableTypeOld", SqlDbType.VarChar).Value = OldTableType;


            cmd.CommandText = @"
                update TablesType set 
                TableType=@TableType,
                SerialNumber=@SerialNumber
                where SN=@SN
                ";
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "更新類別資料";
                //更新管理
                cmd.CommandText = @"
                UPDATE Tables
                   SET TableType = @TableType
                 WHERE  TableType=@TableTypeOld
                ";
                string Result2 = GetSQLNonQuery(cmd);
                if (Result2 == "ok")
                {
                    Result += "、管理資料";
                }
                Result += "成功";

            }
            else
            {
                Result = "更新類別資料失敗";
            }
            return Result;
        }
        public string DeleteTableType(int SN, string TableType,string SerialNumber)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                delete TablesType 
                where SN=@SN
                ";
            cmd.Parameters.Add("@SN", SqlDbType.Int).Value = SN;
            string Result = GetSQLNonQuery(cmd);
            //
            cmd = new SqlCommand();
            cmd.CommandText = @"
                select FAB,TableID 
                from Tables 
                where TableType=@TableType
                ";
            cmd.Parameters.Add("@TableType", SqlDbType.VarChar).Value = TableType;
            DataTable dt = GetSQLDataTable(cmd);
            foreach (DataRow dr in dt.Rows)
            {
                Tables tables = new Tables();
                tables.FAB = dr["FAB"].ToString();
                tables.TableID = dr["TableID"].ToString();
                DeleteTable(tables);
            }
            if (Result == "ok")
            {
                Result = "刪除類別" + SN + ":" + TableType + "成功";
            }
            else
            {
                Result = "刪除類別" + SN + ":" + TableType + "失敗";
            }
            return Result;
        }
        #endregion

        #region TagCode使用表單
        public ActionResult TagCodeManage()
        {
            return View();
        }
        public string TagCodeList(string FAB, int page, int rows)
        {
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            string FABLine = "";
            if (FAB != "ALL")
                FABLine = @"where a.FAB=@FAB";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select * from
                    (select ROW_NUMBER() OVER(ORDER BY a.FAB,a.TableID) AS 'RowNo'
                    ,a.Tag_Code,a.Comment,a.FAB,a.TableID,b.TableName,b.WeeklyCycle,b.MonthCycle,b.YearCycle
                    from TagCodeUseTable as a
	                left join Tables as b
	                on a.TableID = b.TableID
                    {0}
                    ) as t
				where RowNo between @LeftNo and @RightNo
                ", FABLine);
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            if (FAB != "ALL")
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"select COUNT(a.TableID) from TagCodeUseTable as a {0}", FABLine);
            if (FAB != "ALL")
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                string MonthCycle = dr["MonthCycle"].ToString();
                string YearCycle = dr["YearCycle"].ToString();
                string WeeklyCycle = dr["WeeklyCycle"].ToString();
                var i = 0;
                if (MonthCycle == "" || MonthCycle == null) {
                    MonthCycle = "0";
                }
                if (YearCycle == "" || YearCycle == null)
                {
                    YearCycle = "0";
                }
                if (WeeklyCycle == "" || WeeklyCycle == null)
                {
                    WeeklyCycle = "0";
                }
                string[] Array = MonthCycle.Split(',');
                string MonthCycleArray = null;
                //ArrayList MonthCycleArray1 = new ArrayList();//儲存結果的集合
                for (i = 0; i < Array.Length; i++)
                {
                    MonthCycleArray += Convert.ToInt32(Array[i]).ToString() + ",";
                }
                MonthCycleArray = MonthCycleArray.Substring(0, MonthCycleArray.Length - 1);

                Array = YearCycle.Split(',');
                string YearCycleArray = null;
                for (i = 0; i < Array.Length; i++)
                {
                    YearCycleArray += Convert.ToInt32(Array[i]).ToString() + ",";
                }
                YearCycleArray = YearCycleArray.Substring(0, YearCycleArray.Length - 1);

                Array = WeeklyCycle.Split(',');
                string WeeklyCycleArray = null;
                for (i = 0; i < Array.Length; i++)
                {
                    WeeklyCycleArray += Convert.ToInt32(Array[i]).ToString() + ",";
                }
                WeeklyCycleArray = WeeklyCycleArray.Substring(0, WeeklyCycleArray.Length - 1);
                JObject jobj = new JObject
                {
                    {"Tag_Code",dr["Tag_Code"].ToString() },
                    {"FAB",dr["FAB"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"TableName",dr["TableName"].ToString() },
                    {"Comment",dr["Comment"].ToString() },
                    {"DateMonth","星期:"+WeeklyCycleArray.ToString()+"<br>天:"+MonthCycleArray.ToString()+"<br>月:"+YearCycleArray.ToString()}
                };
                jar.Add(jobj);
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"Tag_Code","" },
                    {"FAB","" },
                    {"TableID","" },
                    {"TableName","沒有資料" },
                    {"Comment","" }
                };
            }
            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string NewTagCodeProcess(string Tag_Code, string FAB, string TableID, string Comment)
        {
            string result = "";
            DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"insert into TagCodeUseTable
                    ( Tag_Code, FAB, TableID, Comment)
                    values
                    (@Tag_Code,@FAB,@TableID,@Comment)
                    ";
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
                result = "TagCode新增成功";
            else
                result = "TagCode新增失敗";
            return result;
        }
        public string UpdateTagCode(string Tag_Code, string FAB, string TableID, string Comment)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update TagCodeUseTable set 
                FAB=@FAB,TableID=@TableID,Comment=@Comment
                where Tag_Code=@Tag_Code
                ";
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
                Result = "更新TagCode成功";
            else
                Result = "更新TagCode失敗";
            return Result;
        }
        public string DeleteTagCode(string Tag_Code)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"delete TagCodeUseTable where Tag_Code=@Tag_Code";
            cmd.Parameters.Add("@Tag_Code", SqlDbType.VarChar).Value = Tag_Code;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
                Result = "刪除資料" + Tag_Code + "成功";
            else
                Result = "刪除資料" + Tag_Code + "失敗";
            return Result;
        }
        //去除不要的字元
        //public static string TrimBefore(string origin, char c)
        //{

        //    StringBuilder sb = new StringBuilder(origin);

        //    while (sb.Length > 0 && sb[0] == c)
        //    {

        //        sb.Remove(0, 1);
        //    }

        //    return sb.ToString();

        //}
        #endregion

        #region 點檢表單查詢
        public ActionResult DataManage()
        {
            return View();
        }
        /**
         * 檢點未檢點/異常數值
         * **/
        public string Checkpoint(string DateStart, string DateEnd, string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            DateTime dtime = DateTime.Now;

            string[] ArrayStart = new string[] { };
            if (DateStart != "ALL")
            {
                ArrayStart = DateStart.Split('-');
                DateStart = "";
                for (int i = 0; i < ArrayStart.Length; i++)
                {
                    DateStart += ArrayStart[i];
                }
            }
            string[] ArrayEnd = new string[] { };
            if (DateEnd != "ALL")
            {
                ArrayEnd = DateEnd.Split('-');
                DateEnd = "";
                for (int i = 0; i < ArrayEnd.Length; i++)
                {
                    DateEnd += ArrayEnd[i];
                }
            }
            string where = @"where";
            string[] arrayALL = new string[] { FAB, DateStart, DateEnd };
            string[] arrayALLTxT = new string[] { "FAB" + @" = '" + FAB + @"' and ", "AliveTime" + @" >= '" + DateStart + @"' and ", "AliveTime" + @" <= '" + DateEnd + @"' and " };
            if ((DateStart == DateEnd) && (DateStart != "ALL"))
            {
                if (FAB == "ALL")
                {
                    where += @" AliveTime >= DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"'), '') AND AliveTime<DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"') +1, '') and ";
                }
                else
                {
                    where += @" " + arrayALLTxT[0] + @" AliveTime >= DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"'), '') AND AliveTime < DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"') +1, '') and ";
                }
            }
            else
            {
                if (arrayALL.Length == arrayALLTxT.Length)
                {
                    for (int i = 0; i < arrayALL.Length; i++)
                    {
                        if (arrayALL[i] != "ALL")
                        {
                            where += @" " + arrayALLTxT[i];

                        }
                    }
                }

            }


            string[] IsFinished = new string[] { "where IsFinished = '0'", "where IsFinished = '1'", "" };
            string[] IsFinished2 = new string[] { "and IsFinished = '0'", "and IsFinished = '1'", "" };
            //  string[] IsFinishedTxT = new string[] { "yes", "no", "ALL" };
            //使用Json回傳
            string alldata = "";
            string where2 = "";
            if (IsFinished.Length == IsFinished2.Length)
            {
                for (int i = 0; i < IsFinished.Length; i++)
                {
                    if (where == "where") where2 = IsFinished[i];
                    else where2 = where.Substring(0, where.Length - 5) + @" " + IsFinished2[i];
                   
                    cmd = new SqlCommand();
                    cmd.CommandText = string.Format(@"
                    select COUNT(Doc) 
                    from Datas
                    {0}
                    ", where2);
                    alldata += GetSQLScalar(cmd, Sqlconn) + ",";

                }
            }
            //全部完單
            alldata = alldata.Substring(0, alldata.Length - 1);
            string where3 = @"where";
            string[] arrayALLTxT3 = new string[] { "b.FAB" + @" = '" + FAB + @"' and ", "b.AliveTime" + @" >= '" + DateStart + @"' and ", "b.AliveTime" + @" <= '" + DateEnd + @"' and " };
            if ((DateStart == DateEnd) && (DateStart != "ALL" || DateEnd != "ALL"))
            {

                if (FAB == "ALL")
                {
                    //where  b.AliveTime BETWEEN '20200225' AND '20200226'範圍
                    //where datediff(day,  b.AliveTime, '20200225') = 0 當天
                 //   where3 += @" b.AliveTime >= DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"'), '') AND b.AliveTime<DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"') +1, '') and ";
                    where3 += @" datediff(day,  b.AliveTime, '"+ DateStart + @"') = 0  and ";
                }
                else
                {
                    //    where3 += @" " + arrayALLTxT3[0] + @" b.AliveTime >= DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"'), '') AND b.AliveTime < DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"') +1, '') and ";
                    where3 += @" datediff(day,  b.AliveTime, '" + DateStart + @"') = 0  and ";
                }
            }
            else
            {
                if (arrayALL.Length == arrayALLTxT3.Length)
                {
                    for (int i = 0; i < arrayALL.Length; i++)
                    {
                        if (arrayALL[i] != "ALL")
                        {
                            where3 += @" " + arrayALLTxT3[i];

                        }
                    }
                }

            }

            alldata += "," + AbnormalStatisticsSQL(where3);
            //  alldata += AbnormalStatisticsSQL(where2);
            return alldata;
        }
        /**
         * 完單異常
         * 0.全部
         * 1.正常
         * **/
        public string AbnormalStatisticsSQL(string Where) {
            string Where1 = "";
            string alldata = "";
            string err = "";
            string ann = "";
            //全部欄位 246
            //null 21
            //所有1的欄位194 190=ok  4個不是ok
            //所有2的欄位23  7個ok  1個正確 15個錯誤
            //所有3的欄位註解 8


            SqlCommand cmd = new SqlCommand();
            //全部欄位
            if (Where != "where")
            {
                Where1 = Where.Substring(0, Where.Length - 5);
            }
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                   select count(*) as v from(
	                    select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
	                    from [iCheck].[dbo].DatasItem as a
	                    left join [iCheck].[dbo].Datas as b on a.Doc = b.Doc
	                    left join [iCheck].[dbo].TablesItem as c on b.TableID = c.TableID 
	                    and a.ItemID=c.ItemID	                    
                        {0} 
                    ) as a where a.ItemType is not null ", Where1);
            alldata = GetSQLScalar(cmd, Sqlconn);

            //錯誤的欄位
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                   select count(*) as v from(
	                    select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
	                    from [iCheck].[dbo].DatasItem as a
	                    left join [iCheck].[dbo].Datas as b
	                    on a.Doc = b.Doc
	                    left join [iCheck].[dbo].TablesItem as c
	                    on b.TableID = c.TableID 
	                    and a.ItemID=c.ItemID	                    
                        {0} 
                        c.ItemType ='1' AND a.ItemValue !='ok' 
                    ) as a ", Where);
            err = GetSQLScalar(cmd, Sqlconn);
            //錯誤的欄位
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
            select count(*) as v from(
	            select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
	            from [iCheck].[dbo].DatasItem as a
	            left join [iCheck].[dbo].Datas as b
	            on a.Doc = b.Doc
	            left join [iCheck].[dbo].TablesItem as c
	            on b.TableID = c.TableID 
	            and a.ItemID=c.ItemID
	            {0} 
                (c.ItemType='2') AND a.ItemValue not like '[^-0-9]%'
	
            ) as a
            where CAST(a.ItemValue AS float) < a.ItemMin OR CAST(a.ItemValue AS float) > a.ItemMax

            ", Where);
            err = (Int32.Parse(GetSQLScalar(cmd, Sqlconn)) + Int32.Parse(err)).ToString();
            cmd = new SqlCommand();
            //備註的欄位
            cmd.CommandText = string.Format(@"
                   select count(*) as v from(
	                    select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
	                    from [iCheck].[dbo].DatasItem as a
	                    left join [iCheck].[dbo].Datas as b
	                    on a.Doc = b.Doc
	                    left join [iCheck].[dbo].TablesItem as c
	                    on b.TableID = c.TableID 
	                    and a.ItemID=c.ItemID	                    
                        {0} 
                        (c.ItemType !='2') AND (c.ItemType !='1')
                    ) as a ", Where);
            ann = GetSQLScalar(cmd, Sqlconn);
            alldata = alldata + "," + (Int32.Parse(alldata)  - Int32.Parse(err)).ToString() + "," + err + "," + ann;
            //1.全部2.正確3.錯誤4.註解
            return alldata;

        }
        /**
         * where條件
         * 
         * **/
        private string DataWhere(string[] arrayALL, string[] arrayALLTxT, string DateStart, string DateEnd)
        {
            string where = @"where ";
            string[] ArrayStart = new string[] { };
            if (DateStart != "ALL")
            {
                ArrayStart = DateStart.Split('-');
                DateStart = "";
                for (int i = 0; i < ArrayStart.Length; i++)
                {
                    DateStart += ArrayStart[i];
                }
            }
            string[] ArrayEnd = new string[] { };
            if (DateEnd != "ALL")
            {
                ArrayEnd = DateEnd.Split('-');
                DateEnd = "";
                for (int i = 0; i < ArrayEnd.Length; i++)
                {
                    DateEnd += ArrayEnd[i];
                }
            }
            if ((DateStart == DateEnd) && (DateStart != "ALL"))
            {
                if (arrayALL.Length == arrayALLTxT.Length)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        if (arrayALL[i] != "ALL")
                        {
                            //where += @" " + arrayALLTxT[i] + @" a.AliveTime" + @" = '" + DateStart + @"' and ";
                            where += @" " + arrayALLTxT[i];

                        }
                    }
                    where += @"  AliveTime >= DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"'), '') AND AliveTime<DATEADD(day, DATEDIFF(day, '', '" + DateStart + @"') +1, '') and ";
                }
            }
            else
            {
                if (arrayALL.Length == arrayALLTxT.Length)
                {
                    for (int i = 0; i < arrayALL.Length; i++)
                    {
                        if (arrayALL[i] != "ALL")
                        {
                            where += @" " + arrayALLTxT[i];

                        }
                    }
                }

            }
            if (where == "where ")
            {
                where = @"";
            }
            else
            {
                where = where.Substring(0, where.Length - 5);
            }

            return where;
        }
        /**
         * 傳回網頁的值
         * **/
        public string GetDataList(string DateStart, string DateEnd, string FAB, string IsFinished,string Error, int page, int rows)
        {

            //RowData
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            string From = @"
            from DatasItem as a
		    left join Datas as b        on a.Doc = b.Doc
		    left join TablesItem as c   on b.TableID = c.TableID  and a.ItemID=c.ItemID";
           // left join Tables as d on b.TableID=d.TableID　
           // left join TablesType as e on d.TableType=e.TableType
            
            string[] arrayALL = new string[] { IsFinished, FAB, DateStart, DateEnd };
            string[] arrayALLTxT = new string[] { "a.IsFinished = '" + IsFinished + "' and ", "a.FAB" + @" = '" + FAB + @"' and ", "a.AliveTime" + @" >= '" + DateStart + @"' and ", "a.AliveTime" + @" <= '" + DateEnd + @"' and " };
            string where = DataWhere(arrayALL, arrayALLTxT, DateStart, DateEnd);
            SqlCommand cmd = new SqlCommand();
            //友直的
            if (Error == "ALL") {
                cmd.CommandText = string.Format($@"
                    SELECT DISTINCT a.Doc as Doc from Datas as a 
                    {where}                
                    ORDER BY a.Doc ASC
                ");
            }
            else if (Error=="0")
            {
                cmd.CommandText = string.Format($@"
                    SELECT DISTINCT a.Doc as Doc from(
		                    select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
		                    {From}						
		                    where (c.ItemType ='1' and a.ItemValue='ok') OR 
                            (
                                (c.ItemType='2') AND a.ItemValue ='ok' OR	
                                (c.ItemType='2') AND a.ItemValue not like '[^-0-9]%' AND 
                                CAST(a.ItemValue AS float) >= c.ItemMin AND CAST(a.ItemValue AS float) <= c.ItemMax
                            )
                    ) as a	
                    {where}                        
                    ORDER BY a.Doc ASC
                ");
            }
            else if (Error == "1")
            {
                cmd.CommandText = string.Format($@"
                SELECT DISTINCT a.Doc as Doc from(
	                select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
	                {From}					
	                where (c.ItemType ='1' AND a.ItemValue !='ok') OR 
                    (
		                c.ItemType='2' AND a.ItemValue not like '[^-0-9]%' AND 
		                (
			                CAST(a.ItemValue AS float) < c.ItemMin OR 
			                CAST(a.ItemValue AS float) > c.ItemMax
		                )
	                )
                ) as a	
                {where}                     
                ORDER BY a.Doc ASC
                ");
            }
            else if (Error == "3")
            {
                cmd.CommandText = string.Format($@"
                select DISTINCT a.Doc as Doc from(
	                select b.FAB,a.Doc,a.ItemID,b.IsFinished,a.ItemValue,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax,c.CreateTime,b.AliveTime
	                {From}
	                where c.ItemType is  null 
                ) as a
                {where}   

                ORDER BY a.Doc ASC
                ");
            }
            DataTable dt = GetSQLDataTable(cmd);
            
            string Doc = "";
            string all =dt.Rows.Count.ToString()+","+dt.Columns.Count.ToString();
            foreach (DataRow dr in dt.Rows)
            {
                Doc += "'"+dr["Doc"].ToString()+"',";
            }
            if(Doc != ""){
                if (where != "") Doc = "and a.Doc IN(" + Doc.Substring(0, Doc.Length - 1).ToString() + ")";
                else  Doc = "where a.Doc IN(" + Doc.Substring(0, Doc.Length - 1).ToString() + ")";
            }
            cmd.CommandText = string.Format($@"
            select * from
                (select ROW_NUMBER() OVER(ORDER BY a.AliveTime desc,a.FAB,a.Doc) AS 'RowNo'
                ,a.FAB,a.Doc,a.TableID,t.TableName,a.AliveTime,a.IsFinished,a.IsFinishedTime,a.UserID,e.SerialNumber
                from Datas as a
	            left join Tables as t  on a.TableID=t.TableID
            left join Tables as d on a.TableID=d.TableID　
            left join TablesType as e on d.TableType=e.TableType
                {where} {Doc}
                ) as t
			where RowNo between @LeftNo and @RightNo
            ");
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            //TotalRow.Count
            dt = GetSQLDataTable(cmd);
            //string TotalRow = dt.Rows.Count.ToString();

            where = @"where ";
            arrayALL = new string[] { IsFinished, FAB, DateStart, DateEnd };
            arrayALLTxT = new string[] { "IsFinished = '" + IsFinished + "' and ", "Fab" + @" = '" + FAB + @"' and ", "alivetime" + @" >= '" + DateStart + @"' and ", "alivetime" + @" <= '" + DateEnd + @"' and " };
            where = DataWhere(arrayALL, arrayALLTxT, DateStart, DateEnd);
            //異常/不異常
            if (Error == "ALL")
            {
                cmd = new SqlCommand();
                cmd.CommandText = string.Format($@"
                select count(doc) from datas {where}");
            }
            else if (Error == "0")
            {
                cmd.CommandText = string.Format($@"
                    SELECT COUNT(DISTINCT a.Doc) from(
		                select b.* {From}						
		                where (c.ItemType ='1' and a.ItemValue='ok') OR 
                        (
                            (c.ItemType='2') AND a.ItemValue ='ok' OR	
                            (c.ItemType='2') AND a.ItemValue not like '[^-0-9]%' AND 
                            CAST(a.ItemValue AS float) >= c.ItemMin AND CAST(a.ItemValue AS float) <= c.ItemMax
                        )
                    ) as a	
                    {where}               
                ");
            }
            else if (Error == "1")
            {
                cmd.CommandText = string.Format($@"
                SELECT COUNT(DISTINCT a.Doc) from(
	                select b.* {From}						
	                where (c.ItemType ='1' AND a.ItemValue !='ok') OR 
                    (
		                c.ItemType='2' AND a.ItemValue not like '[^-0-9]%' AND 
		                (
			                CAST(a.ItemValue AS float) < c.ItemMin OR  CAST(a.ItemValue AS float) > c.ItemMax
		                )
	                )
                ) as a	
                {where}                     
                ");
            }
            else if (Error == "3")
            {
                cmd.CommandText = string.Format($@"
                select COUNT(DISTINCT a.Doc) from(
	                select b.* {From}  where c.ItemType is  null 
                ) as a
                {where}   
                ");
            }



            ////使用Json回傳
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            DateTime dtime = DateTime.Now;
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {

                DateTime AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                string AliveTimeLine = AliveTime.ToString("yyyy-MM-dd HH:mm:ss");
                //if (dtime > AliveTime && IsFinished == "0")
                if (dtime > AliveTime && dr["IsFinished"].ToString() == "False")
                    AliveTimeLine = "<span style='color:red'>" + AliveTimeLine + "</span>";

                JObject jobj = new JObject
                {
                    {"FAB",dr["FAB"].ToString() },
                    {"Doc",dr["Doc"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"TableName",dr["TableName"].ToString() },
                     {"SerialNumber",dr["SerialNumber"].ToString() },
                    {"AliveTime",AliveTimeLine },
                    {"IsFinished",dr["IsFinished"].ToString() },
                    {"IsFinishedTime",dr["IsFinishedTime"].ToString() },
                    {"UserID",dr["UserID"].ToString() }
                };
                jar.Add(jobj);
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"FAB","" },
                    {"Doc","" },
                    {"TableID","" },
                    {"SerialNumber","" },
                    {"TableName","沒有資料" },
                      
                    {"AliveTime","" },
                    {"IsFinished","" },
                    {"IsFinishedTime","" },
                    {"UserID","" }
                };
                jar.Add(jobj);
            }


            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        /**
         * 回傳給網頁未檢點值
         * **/
        public string AllUserData(string DateStart, string DateEnd, string FAB) {
            JObject jo = new JObject();
            JArray jar = new JArray();
            string checkpoint = Checkpoint(DateStart, DateEnd, FAB);
            //string abnormalstatistics = AbnormalStatistics(DateStart, DateEnd, FAB);
            checkpoint = checkpoint.Substring(0, checkpoint.Length - 1);
            string[] checkpointData = checkpoint.Split(',');
            for (int i = 0; i < checkpointData.Length; i++) {
                jar.Add(checkpointData[i]);
            }
            jo.Add("ALL", jar);
            return JsonConvert.SerializeObject(jo).ToString();

        }
        
        public string GetTableIDbs(string FAB)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select TableID,TableName
                from Tables
                where FAB=@FAB
                order by TableID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value='{0}' selected>{1}</option>
                        ", dr["TableID"].ToString(), dr["TableName"].ToString());
                    i++;
                }
                else
                {
                    result += string.Format(@"
                        <option value='{0}'>{1}</option>
                        ", dr["TableID"].ToString(), dr["TableName"].ToString());
                }
            }
            if (dt.Rows.Count == 0)
                result += @"<option value='' selected>無資料</option>";
            return result;
        }
        public string CheckDocDuplicate(string Doc)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "select Doc from Datas where Doc=@Doc";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            string Test = GetSQLScalar(cmd, Sqlconn);
            if (Test != "")
                return "t";
            else
                return "f";
        }
        public string AddData(Datas datas)
        {

            string result = "";
            //判斷是不是建過單了
            string SearchDoc = datas.FAB + "_" + datas.TableID + "_" + DateTime.Now.ToString("yyyyMMdd") + "%";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select Doc from Datas
                where Doc like @Doc
                ");
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = SearchDoc;
            string TestDoc = GetSQLScalar(cmd, Sqlconn);
            if (TestDoc != "") return TestDoc + "已存在";
            //建立資料
            DateTime dtime = DateTime.Now;
            cmd = new SqlCommand();
            cmd.CommandText = @"
                insert into Datas
                ( Doc, FAB, TableID, AliveTime, IsFinished, CreateDate)
                values
                (@Doc,@FAB,@TableID,@AliveTime,'0',@CreateDate)
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = datas.Doc;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = datas.FAB;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = datas.TableID;
            cmd.Parameters.Add("@AliveTime", SqlDbType.VarChar).Value = datas.AliveTime;
            cmd.Parameters.Add("@CreateDate", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy-MM-dd");
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
            {
                result = "單號新增成功";
            }
            else
            {
                result = "單號新增失敗";
            }
            return result;
        }
        public string GetDatasItem(string FAB, string Doc,string Error)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select a.Doc,a.ItemID,a.ItemValue,b.IsFinished
		                ,c.ItemName,c.ItemType,c.ItemMin,c.ItemMax
                from DatasItem as a
                left join Datas as b
                on a.Doc = b.Doc
                left join TablesItem as c
                on b.TableID = c.TableID 
                and a.ItemID=c.ItemID
                where b.FAB=@FAB 
                and a.Doc=@Doc
                order by c.ItemSort
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                string ItemValue = dr["ItemValue"].ToString();
                if (dr["ItemType"].ToString() == "1" || dr["ItemType"].ToString() == "4")
                {
                    if ((Error == "1" || Error == "ALL") && ItemValue != "OK" && ItemValue != "N/A") {
                        ItemValue = "<span style='color:red'>" + ItemValue + "</span>";
                        jar.Add(GetDatasItemTrue(dr, Error, ItemValue));
                    } else if ((Error == "0" || Error == "ALL") && (ItemValue == "OK" || ItemValue == "N/A")) {
                        jar.Add(GetDatasItemTrue(dr, Error, ItemValue));
                    }
                    
                }
                else if (dr["ItemType"].ToString() == "2")
                {
                    try
                    {
                        float Value = float.Parse(dr["ItemValue"].ToString());
                        float ItemMin = float.Parse(dr["ItemMin"].ToString());
                        float ItemMax = float.Parse(dr["ItemMax"].ToString());
                        if (Error == "1" || Error == "ALL"){
                            if (Value < ItemMin || Value > ItemMax)
                            {
                                ItemValue = "<span style='color:red'>" + ItemValue + "</span>";
                                jar.Add(GetDatasItemTrue(dr, Error, ItemValue));
                            }
                        }
                        if (Error == "0" || Error == "ALL") {
                            if (Value < ItemMin || Value > ItemMax)
                            {
                                continue;
                            }
                            else {
                                jar.Add(GetDatasItemTrue(dr, Error, ItemValue));
                            }
                        }

                    }
                    catch
                    {
                        if (Error == "ALL") {
                            ItemValue = "<span style='color:yellow'>" + ItemValue + "</span>";
                            jar.Add(GetDatasItemTrue(dr, Error, ItemValue));
                        }
                       
                    }
                }
                //jar.Add(jobj);
                
            }
            if (dt.Rows.Count == 0)
            {
                JObject jobj = new JObject
                {
                    {"Doc","" },
                    {"ItemID","" },
                    {"ItemName","" },
                    {"ItemValue","沒有資料" },
                    {"IsFinished","" },
                    {"ItemType","" },
                    {"ItemMin",""  },
                    {"ItemMax",""  }
                };
                jar.Add(jobj);
            }

            jo.Add("total", dt.Rows.Count);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        private JObject GetDatasItemTrue(DataRow dr,string error,string ItemValue) {
            JObject jobj = new JObject();
          
            jobj = new JObject
            {
                {"Doc",dr["Doc"].ToString() },
                {"ItemID",dr["ItemID"].ToString() },
                {"ItemName",dr["ItemName"].ToString() },
                {"ItemValue",ItemValue },
                {"IsFinished",dr["IsFinished"].ToString() },
                {"ItemType",dr["ItemType"].ToString() },
                {"ItemMin",dr["ItemMin"].ToString() },
                {"ItemMax",dr["ItemMax"].ToString() }
            };
            
            return jobj;
        }
        #endregion

        #region 點檢表單 測試頁面
        public ActionResult TestPage()
        {
            return View();
        }
        public string GetDocbs(string FAB)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select Doc,TableID,AliveTime,IsFinished
                from Datas
                where FAB=@FAB
                and AliveTime > @dtime
                and IsFinished='0'
                order by Doc
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            DataTable dt = GetSQLDataTable(cmd);
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
        public string GetTestData(string FAB, string Doc, string TableID)
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
            DataTable dt = GetSQLDataTable(cmd);
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
        public string UpdateDatas(string json)
        {
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
                Result = UpdateDatasItem(Doc, ItemID, ItemValue, UserID);
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select AliveTime,IsFinished
                from Datas
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = GetSQLDataTable(cmd);
            //
            DateTime dtime = DateTime.Now;
            DateTime AliveTime = DateTime.Now.AddHours(-1);
            bool IsFinished = true;
            foreach (DataRow dr in dt.Rows)
            {
                AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                IsFinished = bool.Parse(dr["IsFinished"].ToString());
                break;
            }
            //檢查時間是否OK 
            if (dtime > AliveTime) return "單號時間已過期";
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
            cmd.Parameters.Add("@dtime", SqlDbType.DateTime).Value = dtime;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
                Result = "儲存單號" + Doc + "成功";
            else
                Result = "儲存單號" + Doc + "失敗";
            return Result;
        }
        public string UpdateDatasItem(string Doc, string ItemID, string ItemValue, string UserID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select AliveTime,IsFinished
                from Datas
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = GetSQLDataTable(cmd);
            //
            DateTime dtime = DateTime.Now;
            DateTime AliveTime = DateTime.Now.AddHours(-1);
            bool IsFinished = true;
            foreach (DataRow dr in dt.Rows)
            {
                AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                IsFinished = bool.Parse(dr["IsFinished"].ToString());
                break;
            }
            //檢查時間是否OK 
            if (dtime > AliveTime) return "";
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
            cmd.Parameters.Add("@Create_at", SqlDbType.VarChar).Value = dtime.ToString("yyyy-MM-dd HH:mm:ss");
            string Result = GetSQLNonQuery(cmd);
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
            string Result = GetSQLNonQuery(cmd);
        }
        #endregion


        #region 點檢表單 測試頁面 過期表單
        public ActionResult ExpiredPage()
        {
            return View();
        }
        public string GetExpiredDocbs(string FAB,string DateStart,string DateEnd)
        {
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            SqlCommand cmd = new SqlCommand();
            string IsFinished = "0";
            string[] arrayALL = new string[] { IsFinished, FAB, DateStart, DateEnd };
            string[] arrayALLTxT = new string[] { "a.IsFinished = '" + IsFinished + "' and ", "a.FAB" + @" = '" + FAB + @"' and ", "a.AliveTime" + @" >= '" + DateStart + @"' and ", "a.AliveTime" + @" <= '" + DateEnd + @"' and " };
            string where = DataWhere(arrayALL, arrayALLTxT, DateStart, DateEnd);
            cmd.CommandText = string.Format($@"
                select Doc,TableID,AliveTime,IsFinished
                from Datas as a
                {where}");
            //cmd.CommandText = string.Format(@"
            //    select Doc,TableID,AliveTime,IsFinished
            //    from Datas
            //    where FAB=@FAB
            //    and AliveTime > @dtime
            //    and IsFinished='0'
            //    order by Doc
            //    ");
            //cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            //cmd.Parameters.Add("@dtime", SqlDbType.VarChar).Value = dtime;
            //cmd.Parameters.Add("@DateStart", SqlDbType.VarChar).Value = DateStart;
            //cmd.Parameters.Add("@DateEnd", SqlDbType.VarChar).Value = DateEnd;
            DataTable dt = GetSQLDataTable(cmd);
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
        public string GetExpiredData(string FAB, string Doc, string TableID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select *
                from TablesItem
                where TableID=@TableID
                order by ItemSort
                ");
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            DataTable dt = GetSQLDataTable(cmd);
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
                    ChangeExpiredSortGo(TableID, dr["ItemID"].ToString(), i);
                }
                ItemList += dr["ItemID"].ToString() + ",";
                //畫面使用
                string SelectLine = "";
                if (dr["ItemType"].ToString() == "1")
                {
                    SelectLine = string.Format(@"
                        <input type=""text"" class=""form-control"" id=""{0}"" disabled/>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','OK')"">正常</button>
                        </div>
                        <div class=""input-group-append"">
                            <button class=""btn btn-outline-secondary"" type=""button"" onclick=""ChangeOK('{0}','Error')"">異常</button>
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
        public string UpdateExpiredDatas(string json)
        {
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
                Result = UpdateExpiredDatasItem(Doc, ItemID, ItemValue, UserID);
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select AliveTime,IsFinished
                from Datas
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = GetSQLDataTable(cmd);
            //
            DateTime dtime = DateTime.Now;
            DateTime AliveTime = DateTime.Now.AddHours(-1);
            bool IsFinished = true;
            foreach (DataRow dr in dt.Rows)
            {
                AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                IsFinished = bool.Parse(dr["IsFinished"].ToString());
                break;
            }
            //檢查時間是否OK 
            //if (dtime > AliveTime) return "單號時間已過期";
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
            cmd.Parameters.Add("@dtime", SqlDbType.DateTime).Value = AliveTime.ToString("yyyy-MM-dd") + " " + dtime.ToString("HH:mm:ss");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
                Result = "儲存單號" + Doc + "成功";
            else
                Result = "儲存單號" + Doc + "失敗";
            return Result;
        }
        public string UpdateExpiredDatasItem(string Doc, string ItemID, string ItemValue, string UserID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select AliveTime,IsFinished
                from Datas
                where Doc=@Doc
                ";
            cmd.Parameters.Add("@Doc", SqlDbType.VarChar).Value = Doc;
            DataTable dt = GetSQLDataTable(cmd);
            //
            DateTime dtime = DateTime.Now;
            DateTime AliveTime = DateTime.Now.AddHours(-1);
            bool IsFinished = true;
            foreach (DataRow dr in dt.Rows)
            {
                AliveTime = DateTime.Parse(dr["AliveTime"].ToString());
                IsFinished = bool.Parse(dr["IsFinished"].ToString());
                break;
            }
            //檢查時間是否OK 
            //if (dtime > AliveTime) return "";
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
            cmd.Parameters.Add("@Create_at", SqlDbType.VarChar).Value = AliveTime.ToString("yyyy-MM-dd") +" "+dtime.ToString("HH:mm:ss");
            string Result = GetSQLNonQuery(cmd);
            //成功 OK 失敗 ""
            return Result;
        }
        private void ChangeExpiredSortGo(string TableID, string ItemID, int ItemSort)
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
            string Result = GetSQLNonQuery(cmd);
        }
        #endregion




        #region 權限人數管理
        private string[] InfoPeople() {
            int people = 7;
            // int SQLpeople = 1;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                  SELECT COUNT(IsInfo)
                  FROM [iCheck].[dbo].[Users]
                  where IsInfo='true'
                ";
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            string[] array = new string[] { };
            if (people > int.Parse(TotalRow))
            {
               
                array = new string[]{"ok", "符合人數，可新增" + (people - int.Parse(TotalRow)).ToString() + "人" };
               
            }
            else {
                array = new string[] { "error", "不符合人數，請刪除" + (int.Parse(TotalRow) - people).ToString() + "人" };
               
            }
            return array;
        }
        
        #endregion
        #region 帳號管理
        public ActionResult AccountManage()
        {
            return View();
        }
        public string UserList(string FAB, int page, int rows)
        {
            string Admin = (string)Session["LoginAccount"];
            string Common = (string)Session["CommonAccount"];
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            string FABLine = "";
            string admin = "";
            SqlCommand cmd = new SqlCommand();
            if (FAB != "ALL") FABLine = "where FAB like '%'+@FAB+'%'";
            if (Common != null)
            { admin = "where UserID=@uid"; }

            cmd.CommandText = string.Format(@"
                select * from
                    (select ROW_NUMBER() OVER(ORDER BY FAB,IsBoss desc,UserID) AS 'RowNo'
                    ,UserID,UserName,FAB,UserMail,APPIsAdmin,IsAdmin,IsBoss,IsInfo,IsForm
                    from Users
                    {0}{1}) as t
				where RowNo between @LeftNo and @RightNo
                ", FABLine, admin);
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            if (FAB != "ALL")
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            if (Common != null)
            {
                cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = Common;
            }
            DataTable dt = GetSQLDataTable(cmd);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select COUNT(UserID) 
                from Users
                {0}
                ", FABLine);
            if (FAB != "ALL")
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            if (Common != null)
            {
                cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = Common;
            }
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();  
            foreach (DataRow dr in dt.Rows)
            {
                string APPIsAdmin = "無";
                if (dr["APPIsAdmin"].ToString() == "1") APPIsAdmin = "一般使用者";
                else if (dr["APPIsAdmin"].ToString() == "2") APPIsAdmin = "<span style='color:red'>管理者</span>";
                else if (dr["APPIsAdmin"].ToString() == "3") APPIsAdmin = "<span style='color:red'>主管</span>";
                string IsAdmin = "無";
                if (dr["IsAdmin"].ToString() == "1") IsAdmin = "一般使用者";
                else if (dr["IsAdmin"].ToString() == "2") IsAdmin = "<span style='color:red'>管理者</span>";
                string IsBoss = "一般職員";
                if (bool.Parse(dr["IsBoss"].ToString())) IsBoss = "<span style='color:red'>主管<span>";
                string IsInfo = "無";
                if (bool.Parse(dr["IsInfo"].ToString())) IsInfo = "<span style='color:red'>有<span>";
                string IsForm = "無";
                if (dr["IsForm"].ToString()!="") {
                    if (bool.Parse(dr["IsForm"].ToString())) IsForm = "<span style='color:red'>有<span>";
                }
                JObject jobj = new JObject
                {
                    {"UserID",dr["UserID"].ToString() },
                    {"UserName",dr["UserName"].ToString() },
                    {"FAB",dr["FAB"].ToString() },
                    {"UserMail",dr["UserMail"].ToString() },
                    {"APPIsAdmin",dr["APPIsAdmin"].ToString() },
                    {"APPIsAdmin2",APPIsAdmin },
                    {"IsAdmin",dr["IsAdmin"].ToString() },
                    {"IsAdmin2",IsAdmin },
                    {"IsBoss",dr["IsBoss"].ToString() },
                    {"IsBoss2",IsBoss },
                    {"IsInfo",dr["IsInfo"].ToString() },
                    {"IsInfo2",IsInfo },
                    {"IsForm",dr["IsForm"].ToString() },
                    {"IsForm2",IsForm }
                };
                jar.Add(jobj);
            }
            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string CheckAccountDuplicate(string Account)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select UserID 
                from Users 
                where UserID=@UserID
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Account;
            string UserID = GetSQLScalar(cmd, Sqlconn);
            if (UserID != "")
                return "t";
            else
                return "f";
        }
        public string NewAccountProcess(UserData userData)
        {
            string result = "";
            DateTime dtime = DateTime.Now;
            SqlCommand cmd = new SqlCommand();
            string[] InfoPeopleOK = InfoPeople();
            if (InfoPeopleOK[0] == "ok") {
                cmd.CommandText = @"insert into Users
                ( UserID, UserName, UserPwd, FAB, UserMail, APPIsAdmin, IsAdmin, IsBoss,IsInfo,IsForm, CreateTime)
                values
                (@UserID,@UserName,@UserPwd,@FAB,@UserMail,@APPIsAdmin,@IsAdmin,@IsBoss,@IsInfo,@IsForm,@CreateTime)
                ";
                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = userData.UserID;
                cmd.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userData.UserName;
                cmd.Parameters.Add("@UserPwd", SqlDbType.VarChar).Value = userData.UserPwd??"";
                cmd.Parameters.Add("@IsForm", SqlDbType.VarChar).Value = userData.IsForm;
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = userData.FAB;
                cmd.Parameters.Add("@APPIsAdmin", SqlDbType.VarChar).Value = userData.APPIsAdmin;
                cmd.Parameters.Add("@UserMail", SqlDbType.VarChar).Value = userData.UserMail;
                cmd.Parameters.Add("@IsAdmin", SqlDbType.VarChar).Value = userData.IsAdmin;
                cmd.Parameters.Add("@IsBoss", SqlDbType.VarChar).Value = userData.IsBoss;
                cmd.Parameters.Add("@IsInfo", SqlDbType.VarChar).Value = userData.IsInfo;
                cmd.Parameters.Add("@CreateTime", SqlDbType.DateTime).Value = dtime;
                result = GetSQLNonQuery(cmd);
                //*/
                if (result == "ok")
                {
                    result = "帳號新增成功";
                }
                else
                {
                    result = "帳號新增失敗";
                }
            }
            else if (InfoPeopleOK[0] == "error") {

                result = InfoPeopleOK[1];
            }
            else {
                result = "異常";
            }
           
           
            return result;
        }
        public string UpdateEditData(UserData userData)
        {
            SqlCommand cmd = new SqlCommand();
            string[] InfoPeopleOK = InfoPeople();
            if (userData.IsInfo== "False") {
                InfoPeopleOK = new string[] {"ok","跳過判斷是否超過人數限制部分" };
            }
            string Result = "";
            if (InfoPeopleOK[0] == "ok")
            {

                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = userData.UserID;
                cmd.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userData.UserName;
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = userData.FAB;
                cmd.Parameters.Add("@UserMail", SqlDbType.VarChar).Value = userData.UserMail;
                cmd.Parameters.Add("@APPIsAdmin", SqlDbType.VarChar).Value = userData.APPIsAdmin;
                cmd.Parameters.Add("@IsAdmin", SqlDbType.VarChar).Value = userData.IsAdmin;
                cmd.Parameters.Add("@IsBoss", SqlDbType.VarChar).Value = userData.IsBoss;
                cmd.Parameters.Add("@IsInfo", SqlDbType.VarChar).Value = userData.IsInfo;
                cmd.Parameters.Add("@IsForm", SqlDbType.VarChar).Value = userData.IsForm;
                cmd.Parameters.Add("@UserPwd", SqlDbType.VarChar).Value = userData.UserPwd??"";
                if (userData.UserPwd != null && userData.UserPwd != "")
                {
                    cmd.CommandText = @"
                    update users set 
                    UserName=@UserName,FAB=@FAB,UserMail=@UserMail
                    ,APPIsAdmin=@APPIsAdmin,IsAdmin=@IsAdmin,IsBoss=@IsBoss,IsInfo=@IsInfo
                    ,UserPwd=@UserPwd,IsForm=@IsForm 
                    where UserID=@UserID
                    ";

                }
                else {
                    cmd.CommandText = @"
                    update users set 
                    UserName=@UserName,FAB=@FAB,UserMail=@UserMail
                    ,APPIsAdmin=@APPIsAdmin,IsAdmin=@IsAdmin,IsBoss=@IsBoss,IsInfo=@IsInfo,IsForm=@IsForm
                    where UserID=@UserID
                    ";

                }
                Result = GetSQLNonQuery(cmd);
                if (Result == "ok") Result = "更新資料成功";
                else Result = "更新資料失敗";
            }
            else if (InfoPeopleOK[0] == "error")
            {

                Result = InfoPeopleOK[1];
            }
            else
            {
                Result = "異常";
            }
            
            return Result;
        }
        public string DeleteEditData(string UserID)
        {
            string Result = "刪除資料" + UserID + "失敗";
            if (UserID == "" || UserID == null) {
                return Result;
            }
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"SELECT UserID 
                FROM Users 
                where UserID=@UserID";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            DataTable dt = GetSQLDataTable(cmd);
            if (dt.Rows.Count == 0)
            {
                return Result;
            }

            cmd = new SqlCommand();
            cmd.CommandText = @"
                select count(*) as count
                from(	
                SELECT UserID
                FROM AccountAgent 
                where UserID = @UserID
                UNION ALL
                SELECT UserID 
                FROM AccountUseTable
                where UserID = @UserID
                UNION ALL
                SELECT UserID
                FROM  Datas
                where UserID = @UserID
                UNION ALL
                SELECT UserID
                FROM  Tables
                where UserID = @UserID

                ) as a
            ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;

            dt = GetSQLDataTable(cmd);
            //Result += dt.Rows.Count;
            foreach (DataRow dr in dt.Rows) {
                if (dr["count"].ToString() == "0")
                {
                    cmd = new SqlCommand();
                    cmd.CommandText = @"delete users where UserID=@UserID";
                    cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
                    Result = GetSQLNonQuery(cmd);
                    if (Result == "ok")
                    {
                        Result = "刪除資料" + UserID + "成功";
                    }
                }
                else
                {
                    Result += "此" + UserID + "有被使用過，無法進行刪除";
                }

            }

            return Result;
        }
        #endregion

        #region 匯入人員清單
        public string PostExcelUserData()
        {
            string Result = "";
            //Session檢查
            if (Session["LoginAccount"].ToString() == null || Session["LoginAccount"].ToString() == "")
                return "逾時過期請重新登入";
            //取CSV資料
            DataTable Items = new DataTable();
            try
            {
                HttpFileCollection files = System.Web.HttpContext.Current.Request.Files;
                HttpPostedFile file = files[0];
                //HttpPostedFile file = CSVPath;
                string targetDir = System.Web.HttpContext.Current.Server.MapPath("~/Txt");
                string path = Path.Combine(targetDir, Path.GetFileName(file.FileName));
                file.SaveAs(path);
                Items = GetUpdateUserCSV(path);
            }
            catch
            {
                return "檔案格式錯誤";
            }
            //存SQL
            string dtime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            int SaveCount = 0;
            foreach (DataRow dr in Items.Rows)
            {
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = string.Format(@"
                    select UserID from Users
                    where UserID=@UserID
                ");
                cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = dr["UserID"].ToString();
                string TestUserID = GetSQLScalar(cmd, Sqlconn);
                if (TestUserID == "")
                {
                    cmd = new SqlCommand();
                    cmd.CommandText = string.Format(@"
                        insert into Users 
                        ( UserID, UserName, UserPwd, UserMail, FAB, AppIsAdmin, IsAdmin, IsBoss, CreateTime)
                        values
                        (@UserID,@UserName,@UserPwd,@UserMail,@FAB,@AppIsAdmin,@IsAdmin,@IsBoss,@CreateTime)
                    ");
                    cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = dr["UserID"].ToString();
                    cmd.Parameters.Add("@UserName", SqlDbType.VarChar).Value = dr["UserName"].ToString();
                    cmd.Parameters.Add("@UserPwd", SqlDbType.VarChar).Value = dr["UserPwd"].ToString();
                    cmd.Parameters.Add("@UserMail", SqlDbType.VarChar).Value = dr["UserMail"].ToString();
                    cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = dr["FAB"].ToString();
                    cmd.Parameters.Add("@AppIsAdmin", SqlDbType.VarChar).Value = "1";
                    cmd.Parameters.Add("@IsAdmin", SqlDbType.VarChar).Value = "0";
                    bool IsBoss = false;
                    if (dr["JobName"].ToString().Contains("經理")) IsBoss = true;
                    cmd.Parameters.Add("@IsBoss", SqlDbType.Bit).Value = IsBoss;
                    cmd.Parameters.Add("@CreateTime", SqlDbType.VarChar).Value = dtime;
                    string Test = GetSQLNonQuery(cmd);
                    if (Test == "ok") SaveCount++;
                }
                else
                {
                    cmd = new SqlCommand();
                    cmd.CommandText = string.Format(@"
                        Update Users set
                        UserName=@UserName, UserMail=@UserMail, FAB=@FAB,IsBoss=@IsBoss
                        where UserID=@UserID
                    ");
                    cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = dr["UserID"].ToString();
                    cmd.Parameters.Add("@UserName", SqlDbType.VarChar).Value = dr["UserName"].ToString();
                    cmd.Parameters.Add("@UserMail", SqlDbType.VarChar).Value = dr["UserMail"].ToString();
                    cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = dr["FAB"].ToString();
                    bool IsBoss = false;
                    if (dr["JobName"].ToString().Contains("經理")) IsBoss = true;
                    cmd.Parameters.Add("@IsBoss", SqlDbType.Bit).Value = IsBoss;
                    string Test = GetSQLNonQuery(cmd);
                    if (Test == "ok") SaveCount++;
                }
            }
            return Result + "已更新" + SaveCount + "筆資料成功";
        }
        private DataTable GetUpdateUserCSV(string txtpath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("UserID");//0
            dt.Columns.Add("UserName");//1
            dt.Columns.Add("UserPwd");//2
            dt.Columns.Add("UserMail");//3
            dt.Columns.Add("FAB");//5
            dt.Columns.Add("JobName");//6
            string FilePath = txtpath;
            string line;
            using (StreamReader file = new StreamReader(FilePath, Encoding.UTF8))
            {
                while ((line = file.ReadLine()) != null)
                {
                    string[] FileData = line.Split(',');
                    if (FileData[0] == "" || FileData[0] == @"工號") continue;
                    if (FileData[1] == "" || FileData[1] == @"姓名") continue;
                    try
                    {
                        DataRow dr = dt.NewRow();
                        dr["UserID"] = FileData[0].Trim();
                        dr["UserName"] = FileData[1].Trim();
                        dr["UserPwd"] = FileData[2].Trim();
                        dr["UserMail"] = FileData[3].Trim();
                        dr["FAB"] = FileData[5].Trim();
                        dr["JobName"] = FileData[6].Trim();
                        dt.Rows.Add(dr);
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            return dt;
        }
        
        #endregion

        #region 帳號使用表單
        public ActionResult AccountUseTable()
        {

            //ViewData["Common"] = (string)Session["CommonAccount"];
            //ViewData["Admin"] = (string)Session["LoginAccount"];
            return View();
        }
        public string GetAccountList(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct IsBoss,UserID,UserName
                from Users 
                where UserID !=''
                and FAB like '%'+@FAB+'%'
                order by IsBoss desc,UserID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 1)
                {
                    JObject jobj = new JObject
                    {
                        {"Account",dr["UserID"].ToString() },
                        {"AccountName",dr["UserID"].ToString()+" "+dr["UserName"].ToString() },
                        {"selected",true }
                    };
                    jar.Add(jobj);
                }
                else
                {
                    JObject jobj = new JObject
                    {
                        {"Account",dr["UserID"].ToString() },
                        {"AccountName",dr["UserID"].ToString()+" "+dr["UserName"].ToString() }
                    };
                    jar.Add(jobj);
                }
                i++;
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string AccountUseTableList(string Account)
        {
            if (Account == null || Account == "") return "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select a.UserID,a.TableID,a.TableSort
                ,t.FAB,t.TableName,t.TableEnable
                from AccountUseTable as a
                left join Tables as t
                on a.TableID= t.TableID
                where a.UserID=@UserID
                order by t.FAB,a.TableSort
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Account;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                string EnableWord = "<span style='color:red'>停用</span>";
                if (bool.Parse(dr["TableEnable"].ToString()))
                    EnableWord = "<span style='color:blue'>啟用</span>";
                JObject jobj = new JObject
                {
                    {"UserID",dr["UserID"].ToString() },
                    {"FAB",dr["FAB"].ToString() },
                    {"TableID",dr["TableID"].ToString() },
                    {"TableSort",dr["TableSort"].ToString() },
                    {"TableName",dr["TableName"].ToString() },
                    {"TableEnable",dr["TableEnable"].ToString() },
                    {"EnableWord",EnableWord }
                };
                jar.Add(jobj);
            }
            jo.Add("total", dt.Rows.Count);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string GetTableIDList(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select TableID,TableName
                from Tables 
                where FAB=@FAB 
                order by TableID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            string result = "";
            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 0)
                {
                    result += string.Format(@"
                        <option value='{0}' selected>{1}</option>
                        ", dr["TableID"].ToString(), dr["TableName"].ToString());
                    i++;
                }
                else
                {
                    result += string.Format(@"
                        <option value='{0}'>{1}</option>
                        ", dr["TableID"].ToString(), dr["TableName"].ToString());
                }
            }
            if (dt.Rows.Count == 0)
            {
                result += string.Format(@"
                        <option value='' selected>查無表單</option>
                        ");
                i++;
            }
            return result;
        }
        public string NewAccountUseTable(string UserID, string TableID, string TableSort)
        {
            if (UserID == "" || UserID == null || TableID == "" || TableID == null
                 || TableSort == "" || TableSort == null) return "資料不得為空值";
            string result = ""; string TestID = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select UserID
                from AccountUseTable
                where UserID=@UserID
                and TableID=@TableID
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            TestID = GetSQLScalar(cmd, Sqlconn);
            if (TestID != "") return "表單已存在";
            cmd = new SqlCommand();
            cmd.CommandText = @"
                insert into AccountUseTable
                ( UserID, TableID, TableSort)
                values
                (@UserID,@TableID,@TableSort)
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@TableSort", SqlDbType.VarChar).Value = TableSort;
            result = GetSQLNonQuery(cmd);
            if (result == "ok")
            {
                result = "表單新增成功";
            }
            else
            {
                result = "表單新增失敗";
            }
            return result;
        }
        public string UpdateAccountUseTable(string UserID, string TableID, string TableSort)
        {
            if (UserID == "" || UserID == null || TableID == "" || TableID == null
                 || TableSort == "" || TableSort == null) return "資料不得為空值";
            string result = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                Update AccountUseTable
                set TableSort=@TableSort
                where UserID=@UserID
                and TableID=@TableID
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;
            cmd.Parameters.Add("@TableSort", SqlDbType.VarChar).Value = TableSort;
            result = GetSQLNonQuery(cmd);
            if (result == "ok")
            {
                result = "表單更新成功";
            }
            else
            {
                result = "表單更新失敗";
            }
            return result;
        }
        public string DeleteAccountUseTable(string UserID, string TableID)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                delete AccountUseTable
                where UserID=@UserID
                and TableID=@TableID
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@TableID", SqlDbType.VarChar).Value = TableID;

            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "刪除使用表單" + TableID + "成功";
            }
            else
            {
                Result = "刪除使用表單" + TableID + "失敗";
            }
            return Result;
        }
        #endregion

        #region 帳號代理人
        public ActionResult AccountAgent()
        {
            return View();
        }
        public string GetUserList(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct IsBoss,UserID,UserName
                from Users 
                where UserID !=''
                and FAB like '%'+@FAB+'%'
                order by IsBoss desc,UserID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            int i = 1;
            foreach (DataRow dr in dt.Rows)
            {
                if (i == 1)
                {
                    JObject jobj = new JObject
                    {
                        {"Account",dr["UserID"].ToString() },
                        {"AccountName",dr["UserID"].ToString()+" "+dr["UserName"].ToString() },
                        {"selected",true }
                    };
                    jar.Add(jobj);
                }
                else
                {
                    JObject jobj = new JObject
                    {
                        {"Account",dr["UserID"].ToString() },
                        {"AccountName",dr["UserID"].ToString()+" "+dr["UserName"].ToString() }
                    };
                    jar.Add(jobj);
                }
                i++;
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string AccountAgentList(string Account)
        {
            if (Account == null || Account == "") return "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select a.SN,a.UserID,a.Agent_UserID
                    ,b.FAB as Agent_FAB
                    ,b.UserName as Agent_UserName
                from AccountAgent as a
                left join Users as b
                on a.Agent_UserID=b.UserID
                where a.UserID=@UserID
                ");
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = Account;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                string Agent_FAB = dr["Agent_FAB"].ToString();
                string Agent_UserName = dr["Agent_UserName"].ToString();
                if (dr["Agent_UserID"].ToString().Length == 4)
                {
                    if (IsFABCheck(dr["Agent_UserID"].ToString()))
                    {
                        Agent_FAB = dr["Agent_UserID"].ToString();
                        Agent_UserName = dr["Agent_UserID"].ToString() + " 全廠";
                    }
                }
                JObject jobj = new JObject
                {
                    {"SN",dr["SN"].ToString() },
                    {"UserID",dr["UserID"].ToString() },
                    {"Agent_FAB",Agent_FAB },
                    {"Agent_UserID",dr["Agent_UserID"].ToString() },
                    {"Agent_UserName",Agent_UserName }
                };
                jar.Add(jobj);
            }
            jo.Add("total", dt.Rows.Count);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string GetAgentUserList(string FAB)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select distinct IsBoss,UserID,UserName
                from Users 
                where UserID !=''
                and FAB like '%'+@FAB+'%'
                order by IsBoss desc,UserID
                ");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            string result = "";
            //int i = 0;
            if (dt.Rows.Count == 0)
                result += string.Format(@"<option value='' selected>查無表單</option>");
            else
                result += string.Format(@"<option value='{0}' selected>{0} 全廠</option>", FAB);
            foreach (DataRow dr in dt.Rows)
            {
                result += string.Format(@"
                    <option value='{0}'>{1}</option>
                    ", dr["UserID"].ToString(), dr["UserID"].ToString() + " " + dr["UserName"].ToString());
            }
            return result;
        }
        public string NewAccountAgent(string UserID, string Agent_UserID)
        {
            if (UserID == "" || UserID == null || Agent_UserID == "" || Agent_UserID == null)
                return "資料不得為空值";
            if (UserID == Agent_UserID)
                return "代理人不可以是自己";
            string result = ""; string TestID = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                select Agent_UserID
                from AccountUseTable
                where UserID=@UserID
                and Agent_UserID=@Agent_UserID
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Agent_UserID", SqlDbType.VarChar).Value = Agent_UserID;
            TestID = GetSQLScalar(cmd, Sqlconn);
            if (TestID != "") return TestID + "代理人已存在";
            cmd = new SqlCommand();
            cmd.CommandText = @"
                insert into AccountAgent
                ( UserID, Agent_UserID)
                values
                (@UserID,@Agent_UserID)
                ";
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Agent_UserID", SqlDbType.VarChar).Value = Agent_UserID;
            result = GetSQLNonQuery(cmd);
            if (result == "ok")
            {
                result = "新增表單代理人成功";
            }
            else
            {
                result = "新增表單代理人失敗";
            }
            return result;
        }
        public string UpdateAccountAgent(string SN, string UserID, string Agent_UserID)
        {
            if (SN == "" || SN == null || UserID == "" || UserID == null
                 || Agent_UserID == "" || Agent_UserID == null) return "資料不得為空值";
            if (UserID == Agent_UserID)
                return "代理人不可以是自己";
            string result = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                Update AccountAgent
                set Agent_UserID=@Agent_UserID
                where SN=@SN
                and UserID=@UserID
                ";
            cmd.Parameters.Add("@SN", SqlDbType.VarChar).Value = SN;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@Agent_UserID", SqlDbType.VarChar).Value = Agent_UserID;
            result = GetSQLNonQuery(cmd);
            if (result == "ok")
            {
                result = "代理人更新成功";
            }
            else
            {
                result = "代理人更新失敗";
            }
            return result;
        }
        public string DeleteAccountAgent(string SN)
        {
            if (SN == "" || SN == null) return "資料不得為空值";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                delete AccountAgent
                where SN=@SN
                ";
            cmd.Parameters.Add("@SN", SqlDbType.VarChar).Value = SN;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "刪除序號" + SN + "成功";
            }
            else
            {
                Result = "刪除序號" + SN + "失敗";
            }
            return Result;
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
        #endregion

        #region 信箱設定
        public ActionResult MailSetting()
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"select * from Mail_Setting where MailNo='01'");
            DataTable dt = GetSQLDataTable(cmd);
            foreach (DataRow dr in dt.Rows)
            {
                ViewData["MailAccount"] = dr["MailAccount"].ToString();
                ViewData["MailAccountPwd"] = dr["MailAccountPwd"].ToString();
                ViewData["MailSmtp"] = dr["MailSmtp"].ToString();
                ViewData["MailSmtpPort"] = dr["MailSmtpPort"].ToString();
                ViewData["EnableSsl"] = dr["EnableSsl"].ToString();
                ViewData["SendMailFrom"] = dr["SendMailFrom"].ToString();
                ViewData["SendMailTitle"] = dr["SendMailTitle"].ToString();
                break;
            }
            ViewData["token"] = ConfigurationManager.AppSettings["TaskToken"];
            return View();
        }
        public string UpdateMailSetting(string MailAccount, string MailAccountPwd, string MailSmtp, string MailSmtpPort, string EnableSsl, string SendMailFrom, string SendMailTitle)
        {
            string UserID = Session["LoginAccount"].ToString();
            if (UserID == null || UserID == "") return "逾期過時 請重新登入";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                update Mail_Setting
                set MailAccount=@MailAccount, MailAccountPwd=@MailAccountPwd, MailSmtp=@MailSmtp, 
                    MailSmtpPort=@MailSmtpPort, SendMailFrom=@SendMailFrom, SendMailTitle=@SendMailTitle,
                    EnableSsl=@EnableSsl, UserID=@UserID, Update_at=@time
                where MailNo='01'
                ");
            cmd.Parameters.Add("@MailAccount", SqlDbType.VarChar).Value = MailAccount;
            cmd.Parameters.Add("@MailAccountPwd", SqlDbType.VarChar).Value = MailAccountPwd;
            cmd.Parameters.Add("@MailSmtp", SqlDbType.VarChar).Value = MailSmtp;
            cmd.Parameters.Add("@MailSmtpPort", SqlDbType.VarChar).Value = MailSmtpPort;
            cmd.Parameters.Add("@SendMailFrom", SqlDbType.VarChar).Value = SendMailFrom;
            cmd.Parameters.Add("@SendMailTitle", SqlDbType.VarChar).Value = SendMailTitle;
            cmd.Parameters.Add("@EnableSsl", SqlDbType.VarChar).Value = EnableSsl;
            cmd.Parameters.Add("@UserID", SqlDbType.VarChar).Value = UserID;
            cmd.Parameters.Add("@time", SqlDbType.VarChar).Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string result = GetSQLNonQuery(cmd);
            if (result == "ok") result = "更新信箱設定成功";
            return result;
        }
        public string SendTestMail(string TestMail)
        {
            List<string> MailList = new List<string>();
            MailList.Add(TestMail);
            string Subject = "測試信件";
            string Body = "這是由DryPump發出的測試信件";
            MailController.SendMailByGmail(MailList, Subject, Body);
            return "已寄出測試信件";
        }
        #endregion

        #region 假日設定
        public ActionResult HolidaySetting()
        {
            return View();
        }
        public string HolidayList(int page, int rows)
        {
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select * from
                    (select ROW_NUMBER() OVER(ORDER BY Holiday) AS 'RowNo'
                    ,SN,Holiday,Comment
                    from HolidayList
                    ) as t
				where RowNo between @LeftNo and @RightNo
                ");
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            DataTable dt = GetSQLDataTable(cmd);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select COUNT(SN) 
                from HolidayList
                ");
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                DateTime dtime = DateTime.Parse(dr["Holiday"].ToString());
                JObject jobj = new JObject
                {
                    {"SN",dr["SN"].ToString() },
                    {"Holiday",dtime.ToString("yyyy-MM-dd") },
                    {"Comment",dr["Comment"].ToString() }
                };
                jar.Add(jobj);
            }
            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string NewHolidayProcess(string Holiday, string Comment)
        {
            DateTime dtime;
            try {
                dtime = DateTime.Parse(Holiday);
            }
            catch {
                return "日期格式不對";
            }
            if (Comment == null) Comment = "";
            string result = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                insert into HolidayList
                ( Holiday, Comment)
                values
                (@Holiday,@Comment)
                ";
            cmd.Parameters.Add("@Holiday", SqlDbType.Date).Value = dtime;
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
            {
                result = "假日新增成功";
            }
            else
            {
                result = "假日新增失敗";
            }
            return result;
        }
        public string UpdateHoliday(string SN, string Holiday, string Comment)
        {
            DateTime dtime;
            try
            {
                dtime = DateTime.Parse(Holiday);
            }
            catch
            {
                return "日期格式不對";
            }
            if (Comment == null) Comment = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update HolidayList set 
                Holiday=@Holiday,Comment=@Comment
                where SN=@SN
                ";
            cmd.Parameters.Add("@SN", SqlDbType.VarChar).Value = SN;
            cmd.Parameters.Add("@Holiday", SqlDbType.Date).Value = dtime;
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "更新假日成功";
            }
            else
            {
                Result = "更新假日失敗";
            }
            return Result;
        }
        public string DeleteHoliday(string SN)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"delete HolidayList where SN=@SN";
            cmd.Parameters.Add("@SN", SqlDbType.VarChar).Value = SN;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "刪除假日成功";
            }
            else
            {
                Result = "刪除假日失敗";
            }
            return Result;
        }
        #endregion

        #region 留言設定
        public ActionResult MessageSetting()
        {
            return View();
        }
        public string MessageList(string FAB, int page, int rows)
        {
            int LeftNo = (page - 1) * rows + 1;
            int RightNo = page * rows;
            string FABLine = @"where FAB=@FAB";
            if (FAB == "ALL") FABLine = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select * from
                    (select ROW_NUMBER() OVER(ORDER BY StartTime desc,EndTime desc) AS 'RowNo'
                    ,SN,FAB,Comment,StartTime,EndTime
                    from MessageList
	                {0}
                    ) as t
				where RowNo between @LeftNo and @RightNo
                ", FABLine);
            cmd.Parameters.Add("@LeftNo", SqlDbType.Int).Value = LeftNo;
            cmd.Parameters.Add("@RightNo", SqlDbType.Int).Value = RightNo;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = GetSQLDataTable(cmd);
            cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select COUNT(SN) 
                from MessageList
                ");
            string TotalRow = GetSQLScalar(cmd, Sqlconn);
            //使用Json回傳
            JObject jo = new JObject();
            JArray jar = new JArray();
            foreach (DataRow dr in dt.Rows)
            {
                DateTime StartTime = DateTime.Parse(dr["StartTime"].ToString());
                DateTime EndTime = DateTime.Parse(dr["EndTime"].ToString());
                JObject jobj = new JObject
                {
                    {"SN",dr["SN"].ToString() },
                    {"FAB",dr["FAB"].ToString() },
                    {"Comment",dr["Comment"].ToString() },
                    {"StartTime",StartTime.ToString("yyyy-MM-dd") },
                    {"EndTime",EndTime.ToString("yyyy-MM-dd") }
                };
                jar.Add(jobj);
            }
            jo.Add("total", TotalRow);
            jo.Add("rows", jar);
            return JsonConvert.SerializeObject(jo).ToString();
        }
        public string NewMessageProcess(string FAB, string Comment, string StartTime, string EndTime)
        {
            DateTime StartTime_dt;
            DateTime EndTime_dt;
            try
            {
                StartTime_dt = DateTime.Parse(StartTime);
                EndTime_dt = DateTime.Parse(EndTime);
            }
            catch
            {
                return "日期格式不對";
            }
            if (Comment == null) Comment = "";
            string result = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                insert into MessageList
                ( FAB, Comment, StartTime, EndTime)
                values
                (@FAB,@Comment,@StartTime,@EndTime)
                ";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
            cmd.Parameters.Add("@StartTime", SqlDbType.VarChar).Value = StartTime + " 00:00:00";
            cmd.Parameters.Add("@EndTime", SqlDbType.VarChar).Value = EndTime + " 23:59:59";
            result = GetSQLNonQuery(cmd);
            //*/
            if (result == "ok")
            {
                result = "留言新增成功";
            }
            else
            {
                result = "留言新增失敗";
            }
            return result;
        }
        public string UpdateMessage(string SN, string FAB, string Comment, string StartTime, string EndTime)
        {
            DateTime StartTime_dt;
            DateTime EndTime_dt;
            try
            {
                StartTime_dt = DateTime.Parse(StartTime);
                EndTime_dt = DateTime.Parse(EndTime);
            }
            catch
            {
                return "日期格式不對";
            }
            if (Comment == null) Comment = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"
                update MessageList set 
                FAB=@FAB,Comment=@Comment,StartTime=@StartTime,EndTime=@EndTime
                where SN=@SN
                ";
            cmd.Parameters.Add("@SN", SqlDbType.VarChar).Value = SN;
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Comment", SqlDbType.VarChar).Value = Comment;
            cmd.Parameters.Add("@StartTime", SqlDbType.VarChar).Value = StartTime + " 00:00:00";
            cmd.Parameters.Add("@EndTime", SqlDbType.VarChar).Value = EndTime + " 23:59:59";
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "更新留言成功";
            }
            else
            {
                Result = "更新留言失敗";
            }
            return Result;
        }
        public string DeleteMessage(string SN)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = @"delete MessageList where SN=@SN";
            cmd.Parameters.Add("@SN", SqlDbType.VarChar).Value = SN;
            string Result = GetSQLNonQuery(cmd);
            if (Result == "ok")
            {
                Result = "刪除留言成功";
            }
            else
            {
                Result = "刪除留言失敗";
            }
            return Result;
        }
        #endregion

        #region 測試JSON頁面
        public ActionResult TestJson()
        {
            ViewData["RenewDate"] = RenewDate;
            return View();
        }
        public ActionResult TestJsonM()
        {
            ViewData["RenewDate"] = RenewDate;
            return View();
        }
        #endregion

        #region 抓取選項資料
        public string GetSqlFAB()
        {
            string Admin = (string)Session["LoginAccount"];
            string Common = (string)Session["CommonAccount"];
            JArray jar = new JArray();
            SqlCommand cmd = new SqlCommand();
            DataTable dt = new DataTable();
            if (Common != null) {
                cmd.CommandText = string.Format(@"
                select distinct UserID ,distinct FAB
                from Users 
                where FAB !='' order by FAB
                ");

                //使用Json回傳
                dt = GetSQLDataTable(cmd);
                int i = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    if (i == 1)
                    {
                        JObject jobj = new JObject
                    {
                        {"FAB",dr["FAB"].ToString() },
                        {"selected",true }
                    };
                        jar.Add(jobj);
                    }
                    else
                    {
                        JObject jobj = new JObject
                    {
                        {"FAB",dr["FAB"].ToString() },
                    };
                        jar.Add(jobj);
                    }
                    i++;
                }
            }
            if (Admin != null) {
                cmd.CommandText = string.Format(@"
                select distinct FAB
                from Factories 
                where FAB !='' order by FAB
                ");
                dt = GetSQLDataTable(cmd);
                //使用Json回傳

                int i = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    if (i == 1)
                    {
                        JObject jobj = new JObject
                    {
                        {"FAB",dr["FAB"].ToString() },
                        {"selected",true }
                    };
                        jar.Add(jobj);
                    }
                    else
                    {
                        JObject jobj = new JObject
                    {
                        {"FAB",dr["FAB"].ToString() },
                    };
                        jar.Add(jobj);
                    }
                    i++;
                }
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string GetSqlFAB_All(string FAB)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                    select distinct FAB
                    from Factories 
                    where FAB !='' order by FAB
                    ");
            DataTable dt = GetSQLDataTable(cmd);
            //使用Json回傳
            JArray jar = new JArray();
            JObject jobj;
            if (FAB == "" || FAB == null || FAB == "ALL")
            {
                jobj = new JObject
                {
                    {"FAB","ALL" },
                    {"selected",true }
                };
                jar.Add(jobj);
            }
            else
            {
                jobj = new JObject
                {
                    {"FAB","ALL" },
                };
                jar.Add(jobj);
            }
            foreach (DataRow dr in dt.Rows)
            {
                if (FAB == dr["FAB"].ToString())
                {
                    jobj = new JObject
                    {
                        {"FAB",dr["FAB"].ToString() },
                        {"selected",true }
                    };
                    jar.Add(jobj);
                }
                else
                {
                    jobj = new JObject
                    {
                        {"FAB",dr["FAB"].ToString() }
                    };
                    jar.Add(jobj);
                }
            }
            return JsonConvert.SerializeObject(jar).ToString();
        }
        public string GetSqlFABbs(int a)
        {

            string Admin = (string)Session["LoginAccount"];
            string Common = (string)Session["CommonAccount"];
            SqlCommand cmd = new SqlCommand();
            DataTable dt = new DataTable();
            string result = "";
            if (a == 1)
            {
                result += string.Format(@"<option value='ALL' selected>ALL</option>");
            }
            if (Common != null)
            {
                string admin = "";
                admin = "where UserID=@uid";
                cmd.CommandText = string.Format(@"
                select distinct UserID,distinct FAB
                from Users
                {0}
                ", admin);
                cmd.Parameters.Add("@uid", SqlDbType.VarChar).Value = Common;
                dt = GetSQLDataTable(cmd);
                foreach (DataRow dr in dt.Rows)
                {
                    if (a == 0)
                    {
                        result += string.Format(@"
                        <option value='{0}' selected>{0}</option>
                        ", dr["FAB"].ToString());
                        a++;
                    }
                    else
                    {
                        result += string.Format(@"
                        <option value='{0}'>{0}</option>
                        ", dr["FAB"].ToString());
                    }
                }
            }
            else {
                cmd.CommandText = string.Format(@"
                select distinct FAB
                    from Factories 
                    where FAB !='' 
                    order by FAB
                ");
                dt = GetSQLDataTable(cmd);
                //使用Json回傳

                foreach (DataRow dr in dt.Rows)
                {
                    if (a == 0)
                    {
                        result += string.Format(@"
                        <option value='{0}' selected>{0}</option>
                        ", dr["FAB"].ToString());
                        a++;
                    }
                    else
                    {
                        result += string.Format(@"
                        <option value='{0}'>{0}</option>
                        ", dr["FAB"].ToString());
                    }
                }
            }


            return result;
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
                catch (Exception ex)
                {
                    string e = ex.ToString();
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
            catch (Exception ex)
            {
                string e = ex.ToString();
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
                string e = ex.ToString();
            }
            return dt;
        }
        #endregion




        #region Excel
        private string HeaderProcessing(string[] data,string[] txt){
            string FilePath = "";
            string DataAll = "";
            string[] DataAllArray = new string[] { };
            ArrayList array = new ArrayList();
            for (int i=0;i<data.Length;i++) {
                if (txt[i] == "狀態") {
                    if (data[i] == "1") { DataAll = "異常"; }
                    else if (data[i] == "0") { DataAll = "正常"; }
                    else { DataAll = "全部"; }
                   // array.Add(errorData);
                    FilePath += DataAll;
                }
                if (txt[i] == "時間")
                {
                    if (data[i] == "ALL"){FilePath += "";}
                    else
                    {
                        DataAllArray = data[i].Split('-');
                        DataAll = "";
                        for (int j=0;j< DataAllArray.Length;j++) {
                            DataAll += DataAllArray[j];
                        }
                        FilePath += DataAll + "-";
                    }
                }
                if (txt[i] == "廠別")
                {
                    if (data[i] == "ALL"){FilePath += "";}
                    else {FilePath += data[i];}
                   
                }

            }        
            return FilePath;
        }
        public void ExcelDataManageDownload()
        {

            string FAB = Request.Form["FAB"];
            string IsFinished = Request.Form["IsFinished"];
            string DateStart = Request.Form["DateStart"];
            string DateEnd = Request.Form["DateEnd"];
            string Error = Request.Form["Error"];
            JObject jo = new JObject();
            JArray jar = new JArray();
            string[] arrayALL = new string[] { IsFinished, FAB, DateStart, DateEnd };
            string[] arrayALLTxT = new string[] { "a.IsFinished = '" + IsFinished + "' and ", "a.FAB" + @" = '" + FAB + @"' and ", "a.AliveTime" + @" >= '" + DateStart + @"' and ", "a.AliveTime" + @" <= '" + DateEnd + @"' and " };

            string where = DataWhere(arrayALL, arrayALLTxT, DateStart, DateEnd);
            //string where2 = "";
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"
                select t.*,c.ItemID,c.ItemName,b.ItemValue,c.ItemSort,c.ItemType,c.ItemMax,c.ItemMin

                from
                    (select ROW_NUMBER() OVER(ORDER BY a.AliveTime desc, a.FAB, a.Doc) AS 'RowNo'
                    , a.FAB, a.Doc, a.TableID, t.TableName
                    , a.AliveTime, a.IsFinished, a.IsFinishedTime, a.UserID,c.SerialNumber
                    from Datas as a
                    left join Tables as t on a.TableID = t.TableID
                    left join TablesType as c on t.TableType=c.TableType
                    {0}
                    ) as t

                left join DatasItem as b on t.Doc = b.Doc
                left join TablesItem as c
                on t.TableID = c.TableID and b.ItemID=c.ItemID
               
                ", where);

            DataTable dt = GetSQLDataTable(cmd);
            StringBuilder lines = new StringBuilder();
            //    string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員""";
            string checkpoint=Checkpoint(DateStart, DateEnd, FAB);
            string[] checkAray = checkpoint.Split(',');
            string line = @"=""未檢"",="""+checkAray[0]+ @""",=""以檢 "",="""+checkAray[1]+ @""",=""應檢"",="""+checkAray[2]+@"""";

            lines.AppendLine(line);
            line = line + "<br/>";
            line = @"=""廠別"",="""+FAB+@""",=""日期"",="""+ DateStart+@"到"+ DateEnd + @""",=""完單"",="""+ IsFinished + @""",=""異常"",="""+ Error + @"""";
            lines.AppendLine(line);
            line = line + "<br/>";
            line = @"=""廠別"",=""單號"",=""表單ID"",=""表單名稱"",=""規格表單編號"",=""期限時間"",=""完單時間"",=""完單人員"",=""完單"",=""ItemID"",=""ItemName"",=""ItemValue"",=""error""";
            lines.AppendLine(line);
            line = line + "<br/>";
            DateTime today = DateTime.Now;
            string Url = ConfigurationManager.AppSettings["DataManage"];
            string[] DataExcelHeader = new string[] { FAB,DateStart,DateEnd,Error};
            string[] DataExcelHeaderTxT = new string[] { "廠別", "時間", "時間", "狀態" };
            string HeaderProcessingData = HeaderProcessing(DataExcelHeader,DataExcelHeaderTxT);            
            foreach (DataRow dr in dt.Rows)
            {

               
                string error1 = ErrorTrue(dr);
                if (Error == "1" && error1 == "錯誤")
                {
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableID"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";

                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";

                    line += @"=""" + dr["AliveTime"].ToString() + @""",";
                    line += @"=""" + dr["IsFinishedTime"].ToString() + @""",";
                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["IsFinished"].ToString() + @""",";
                    line += @"=""" + dr["ItemID"].ToString() + @""",";
                    line += @"=""" + dr["ItemName"].ToString() + @""",";
                    line += @"=""" + dr["ItemValue"].ToString() + @""",";
                    line += @"=""" + error1 + @""",";
                    lines.AppendLine(line);
                    line = line + "<br/>";
                   // continue;
                } else if (Error == "0" && error1 == "正確") {
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableID"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";

                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";

                    line += @"=""" + dr["AliveTime"].ToString() + @""",";
                    line += @"=""" + dr["IsFinishedTime"].ToString() + @""",";
                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["IsFinished"].ToString() + @""",";
                    line += @"=""" + dr["ItemID"].ToString() + @""",";
                    line += @"=""" + dr["ItemName"].ToString() + @""",";
                    line += @"=""" + dr["ItemValue"].ToString() + @""",";
                    line += @"=""" + error1 + @""",";
                    lines.AppendLine(line);
                    line = line + "<br/>";
                   // continue;
                }
                else if(Error == "ALL") {
                    line = @"=""" + dr["FAB"].ToString() + @""",";
                    line += @"=""" + dr["DOC"].ToString() + @""",";
                    line += @"=""" + dr["TableID"].ToString() + @""",";
                    line += @"=""" + dr["TableName"].ToString() + @""",";

                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";

                    line += @"=""" + dr["AliveTime"].ToString() + @""",";
                    line += @"=""" + dr["IsFinishedTime"].ToString() + @""",";
                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["IsFinished"].ToString() + @""",";
                    line += @"=""" + dr["ItemID"].ToString() + @""",";
                    line += @"=""" + dr["ItemName"].ToString() + @""",";
                    line += @"=""" + dr["ItemValue"].ToString() + @""",";
                    line += @"=""" + error1 + @""",";
                    lines.AppendLine(line);
                    line = line + "<br/>";
                 //   break;
                }
                
                
            }
            ///
         //   string ok = "1";
            try
            {
                //設定標頭
                Response.AddHeader("Content-disposition", "attachment; filename=\"" + HeaderProcessingData + @".csv" + "\"");
                //設定回傳媒體型別(MIME)
                Response.ContentType = "text/vnd.ms-excel";
                //設定主體內容編碼
                Response.ContentEncoding = Encoding.UTF8;
                //建立StreamWriter，取得Response的OutputStream並設定編碼為UTF8
                StreamWriter sw = new StreamWriter(Response.OutputStream, Encoding.UTF8);
                //寫入資料
                sw.Write(lines.ToString());
                //關閉StreamWriter
                sw.Close();
                //釋放StreamWriter資源
                sw.Dispose();
                //送出Response
                Response.End();

                //HttpResponse ResPonse = System.Web.HttpContext.Current.Response;
                //ResPonse.AddHeader(@"Content-Disposition", "attachment; fileName=" + HeaderProcessingData + @".csv");
                //ResPonse.ContentType = "application/vnd.ms-excel";
                //ResPonse.ContentEncoding = Encoding.UTF8;
                //ResPonse.Write(lines.ToString());
               
                //ResPonse.End();

                //FilePath = Url + HeaderProcessingData + @"_" + today.ToString("yyyyMMdd") + ".csv";
                //if (Directory.Exists(Url))
                //{
                //    System.IO.File.WriteAllText(FilePath, lines.ToString(), Encoding.GetEncoding("UTF-8"));
                //}
                //else
                //{
                //    //新增資料夾
                //    Directory.CreateDirectory(Url);
                //    System.IO.File.WriteAllText(FilePath, lines.ToString(), Encoding.GetEncoding("UTF-8"));
                //}
                //return HeaderProcessingData + @"_" + today.ToString("yyyyMMdd") + ".csv";
            }
            catch (Exception ex)
            {
                SaveLog(ex.ToString());
                //return "error";
            }
        }


        private string ErrorTrue(DataRow dr) {
            string error1 = "";
            string ItemValue = dr["ItemValue"].ToString();
            if (dr["ItemType"].ToString() == "1")
            {

                if (ItemValue != "OK") error1 = "錯誤";
            }
            if (dr["ItemType"].ToString() == "4")
            {

                if (ItemValue != "N/A") error1 = "錯誤";
            }

            else if (dr["ItemType"].ToString() == "2")
            {
                try
                {
                    float Value = float.Parse(dr["ItemValue"].ToString());
                    float ItemMin = float.Parse(dr["ItemMin"].ToString());
                    float ItemMax = float.Parse(dr["ItemMax"].ToString());
                    if (Value < ItemMin || Value > ItemMax)
                    {
                        error1 = "錯誤";
                    }

                }
                catch
                {
                    error1 = "正確";
                }
            }
            else
            {
                error1 = "正確";
            }
            return error1;
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
        private string Nullyes(string nulldata) {
            if (nulldata == null) {
                nulldata = "";

            }
            return nulldata;
        }

        #endregion
    }

}