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
    public class TaskOld
    {
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        /// <summary>
        /// 大綱
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string SendTxt1(string FAB, string TxtDate)
        {
            //DateTime today = DateTime.Now;
            DateTime today = SendYMDController.DateString(TxtDate);
            int HH = int.Parse(DateTime.Now.ToString("HH"));
            //try
            //{
            //    if (TxtDate != null && TxtDate != "")  today = DateTime.Parse(TxtDate);
            //}
            //catch { }
            //假日不出單
            StringBuilder Result = new StringBuilder();
            if (DateSetting.HolidayTest(today) != "") return "今天是假日";
            string Weekly = today.DayOfWeek.ToString("d");//星期2
            if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly;
            //
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                    select a.FAB, a.Doc,b.TableName, a.IsFinishedTime, isnull(a.UserID,b.UserID) as UserID, c.UserName,d.SerialNumber,d.TableType
                    from datas a
                    left join tables b on b.TableID=a.TableID
                    left join users c on c.UserID=isnull(a.UserID,b.UserID)
                    left join TablesType as d on b.TableType=d.TableType
                    where a.AliveTime between @TodayF and @TodayL and a.fab = @FAB
                    "
            };
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 00:00:00";
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 23:59:59";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            //表頭
            StringBuilder lines = new StringBuilder();
            bool csv = false;
            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
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
                    else line += @"="""",";
                    line += @"=""" + dr["UserName"].ToString() + @""",";
                    line += @"=""" + dr["UserID"].ToString() + @""",";
                    line += @"=""" + dr["SerialNumber"].ToString() + @""",";
                    lines.AppendLine(line);
                    line += "<br/>";
                    Result.AppendLine(line);
                }
                csv = true;

            }
            else
            {
                if (HH >= int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
                {
                    line += "<br/>";
                    Result.AppendLine(line);
                    csv = true;

                }

            }

            if (csv)
            {

                string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + today.ToString("yyyyMMdd") + ".csv";
                if (!Directory.Exists(ConfigurationManager.AppSettings["TxtUrl"])) Directory.CreateDirectory(ConfigurationManager.AppSettings["TxtUrl"]);
                //*存文字
                TextCsv.SaveName(FilePath, lines, "UTF-8");
                //  Factories_Complete(FAB);//已匯出
            }
            return Result.ToString();
        }
        /// <summary>
        /// 明細
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string SendTxt2(string FAB, string TxtDate)
        {

            int HH = int.Parse(DateTime.Now.ToString("HH"));
            //DateTime today = DateTime.Now;
            DateTime today = SendYMDController.DateString(TxtDate);
            //假日不出單
            StringBuilder Result = new StringBuilder();
            if (DateSetting.HolidayTest(today) != "") return "今天是假日";
            string Weekly = today.DayOfWeek.ToString("d");//星期2
            if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly;
            //
            SqlCommand cmd = new SqlCommand
            {
                CommandText = @"
                select count(doc) from datas
                where doc like '%'+@Today
                and IsFinished = '1' and IsFinishedTime > @TodayF
                and IsFinishedTime < @TodayL and FAB=@FAB
                "
            };
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
            cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 00:00:00";
            cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 23:59:59";
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            int FinishCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            //if (FinishCount == 0) return "";
            cmd = new SqlCommand
            {
                CommandText = @"
                select count(doc)  from datas
                where doc like '%'+@Today and FAB=@FAB
                "
            };
            cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            int TotalCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            DataTable dt = new DataTable();
            //       if (FinishCount >= TotalCount && TotalCount != 0)
            string returnData = "";
            if (TotalCount == 0) return "應完成件數" + TotalCount + "<br/>";
            else
            {
                returnData = "應完成件數" + TotalCount + "<br/>" + "已完成件數" + FinishCount + "<br/>" + "尚未全部完成" + (TotalCount - FinishCount) + "<br/>";
            }
            if (FinishCount > 0 && TotalCount != 0)  //只要完成件數>0就輸出
            {
                cmd = new SqlCommand
                {
                    CommandText = @"
                    select a.FAB,a.Doc,c.TableName,a.IsFinishedTime,a.UserID,e.UserName,d.ItemName,b.ItemValue,d.ItemType,d.ItemMax,d.ItemMin,f.SerialNumber,f.TableType,d.ItemContent
                    from datas as a
                    left join DatasItem as b	on a.doc = b.doc
                    left join Tables as c		on a.TableID = c.TableID
                    left join TablesItem as d	on a.TableID=d.TableID and b.ItemID=d.ItemID
                    left join Users as e		on a.UserID=e.UserID 
                    left join TablesType as f		on c.TableType=f.TableType
                    where a.doc like '%'+@Today
                    and a.IsFinished = '1'
                    and a.IsFinishedTime > @TodayF
                    and a.IsFinishedTime < @TodayL
                    and a.FAB=@FAB
                    "
                };
                cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
                cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 00:00:00";
                cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 23:59:59";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
                dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            }
            else
            {

            }
            if (dt.Rows.Count != 0)
            {
                //表頭
                StringBuilder lines = new StringBuilder();
                //string line = "廠別,單號,表單名稱,完單時間,完單人員,人員姓名,被代理人工號,被代理人姓名,項目名稱,標準值,查詢結果,資料類別,是否異常,異常原因";
                //string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""規格表單編號""";
                //string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因""";
                string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""項目說明"",=""規格表單編號""";
                lines.AppendLine(line);

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
                //            string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + @"" + FAB + @"\" + FAB + @"_" + today.ToString("yyyyMMdd") + ".csv";
                if (!Directory.Exists(ConfigurationManager.AppSettings["TxtUrl"]))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["TxtUrl"] + "SendTxt2");
                }
                string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + @"" + FAB + @"_" + today.ToString("yyyyMMdd") + "list.csv";
                //*存文字
                TextCsv.SaveName(FilePath, lines, "UTF-8");
            }
            else
            {
                return "無資料匯出";
            }

            return returnData + "<br>" + Result.ToString();
        }

        #region 月明細、大綱
        /// <summary>
        /// 產出報表
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="StartDate"></param>
        /// <param name="EndDate"></param>
        /// <param name="Where"></param>
        /// <param name="WhereJoin"></param>
        /// <param name="FAB"></param>
        /// <param name="DateD"></param>
        /// <returns></returns>
        public static string SendOne(SqlCommand cmd, DateTime StartDate, DateTime EndDate, string Where, string WhereJoin, string FAB, string DateD)
        {
            //今天小時
            int HH = int.Parse(DateTime.Now.ToString("HH"));
            StringBuilder Result = new StringBuilder();
            string FilePathHtml = "";
            string whereData = "";
            try
            {

                if (DateD == "MM" || DateD == "DD")
                {
                    whereData = @"where a.doc like @Today and  a.AliveTime between @TodayF and @TodayL and a.FAB=@FAB ";
                }
                else if (DateD == "All")
                {
                    whereData = " where  a.AliveTime between @TodayF and @TodayL and a.FAB=@FAB ";
                    WhereJoin = "";
                    Where = "";
                }
                cmd.CommandText = String.Format(@"
                    select * from (
                        select a.FAB, a.Doc,b.TableName, a.IsFinishedTime, isnull(a.UserID,b.UserID) as UserID, c.UserName,d.SerialNumber,d.TableType
                        from datas a
                        left join tables b on b.TableID=a.TableID
                        left join users c on c.UserID=isnull(a.UserID,b.UserID)
                        left join TablesType as d on b.TableType=d.TableType
                        {0} {1}
                    ) as a", whereData, WhereJoin);

                DataTable dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
                //表頭
                StringBuilder lines = new StringBuilder();
                bool csv = false;
                string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
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
                            line += @"=""" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @""",";
                        else line += @"="""",";
                        line += @"=""" + dr["UserName"].ToString() + @""",";
                        line += @"=""" + dr["UserID"].ToString() + @""",";
                        line += @"=""" + dr["SerialNumber"].ToString() + @""",";
                        lines.AppendLine(line);
                        line += "<br/>";
                        Result.AppendLine(line);
                    }
                    csv = true;

                }
                else
                {
                    if (HH >= int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
                    {
                        line += "<br/>";
                        Result.AppendLine(line);
                        csv = true;

                    }

                }

                if (csv)
                {
                    string CsvDate = "";
                    string StartDateString = StartDate.ToString("yyyyMMdd");
                    string EndDateString = EndDate.ToString("yyyyMMdd");

                    if (StartDateString == EndDateString)
                    {
                        if (DateD == "MM") CsvDate += StartDate.ToString("yyyyMM") + "MM";
                        else if (DateD == "All") CsvDate += StartDate.ToString("yyyyMM");
                        else if (DateD == "DD") CsvDate += StartDate.ToString("yyyyMMdd");

                        //                        CsvDate = StartDateString;
                    }
                    else
                    {
                        if (DateD == "MM") CsvDate += StartDate.ToString("yyyyMM") + "MM";
                        else if (DateD == "All") CsvDate += StartDate.ToString("yyyyMM");
                        else if (DateD == "DD") CsvDate += StartDate.ToString("yyyyMMdd");

                        //                      CsvDate = StartDate.ToString("yyMMdd") + "到"+ EndDate.ToString("yyMMdd");
                    }
                    string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + CsvDate + ".csv";
                    if (!Directory.Exists(ConfigurationManager.AppSettings["TxtUrl"])) Directory.CreateDirectory(ConfigurationManager.AppSettings["TxtUrl"]);
                    //*存文字
                    System.IO.File.WriteAllText(FilePath, lines.ToString(), Encoding.GetEncoding("UTF-8"));
                    //  Factories_Complete(FAB);//已匯出
                    FilePathHtml = FilePath + "<br>";
                }
                return FilePathHtml;

            }
            catch (Exception ex)
            {
                MSSQL.SaveLog(ex.ToString());
                return "錯誤";
            }

        }


        /// <summary>
        /// 產出報表明細
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="StartDate"></param>
        /// <param name="EndDate"></param>
        /// <param name="Where"></param>
        /// <param name="WhereJoin"></param>
        /// <param name="FAB"></param>
        /// <param name="DateD"></param>
        /// <returns></returns>
        public static string SendOneList(SqlCommand cmd, DateTime StartDate, DateTime EndDate, string Where, string WhereJoin, string FAB, string DateD)
        {
            //今天小時
            int HH = int.Parse(DateTime.Now.ToString("HH"));
            StringBuilder Result = new StringBuilder();
            try
            {
                string FilePathHtml = "已產生csv如下";
                string whereData = "";
                if (DateD == "MM" || DateD == "DD")
                {
                    whereData = @"where a.doc like @Today and  a.IsFinished = '1' and a.AliveTime between @TodayF and @TodayL  and a.FAB=@FAB ";
                }
                else if (DateD == "All")
                {
                    whereData = " where  a.IsFinished = '1' and  a.AliveTime between @TodayF and @TodayL and a.FAB=@FAB ";
                    WhereJoin = "";
                    Where = "";
                }

                DataTable dt = new DataTable();
                //       if (FinishCount >= TotalCount && TotalCount != 0)
                // string returnData = "";
                cmd.CommandText = String.Format(@"
                    select * from (
                        select a.FAB,a.Doc,c.TableName,a.IsFinishedTime,a.UserID,e.UserName,d.ItemName,b.ItemValue,d.ItemType,d.ItemMax,d.ItemMin,f.SerialNumber,f.TableType,d.ItemContent
                        from datas as a
                        left join DatasItem as b	on a.doc = b.doc
                        left join Tables as c		on a.TableID = c.TableID
                        left join TablesItem as d	on a.TableID=d.TableID and b.ItemID=d.ItemID
                        left join Users as e		on a.UserID=e.UserID 
                        left join TablesType as f		on c.TableType=f.TableType
                            
                    {0} {1}) as a 
                    ", whereData, WhereJoin);
                dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);

                if (dt.Rows.Count != 0)
                {
                    //表頭
                    StringBuilder lines = new StringBuilder();
                    //string line = "廠別,單號,表單名稱,完單時間,完單人員,人員姓名,被代理人工號,被代理人姓名,項目名稱,標準值,查詢結果,資料類別,是否異常,異常原因";
                    // string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""規格表單編號""";
                    //string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因""";
                    string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""項目說明"",=""規格表單編號""";
                    lines.AppendLine(line);

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
                    //            string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + @"" + FAB + @"\" + FAB + @"_" + today.ToString("yyyyMMdd") + ".csv";
                    if (!Directory.Exists(ConfigurationManager.AppSettings["TxtUrl"]))
                    {
                        Directory.CreateDirectory(ConfigurationManager.AppSettings["TxtUrl"] + "SendTxt2");
                    }
                    string CsvDate = "";
                    string StartDateString = StartDate.ToString("yyyyMMdd");
                    string EndDateString = EndDate.ToString("yyyyMMdd");

                    if (StartDateString == EndDateString)
                    {
                        if (DateD == "MM") CsvDate += StartDate.ToString("yyyyMM") + "MM_list";
                        else if (DateD == "All") CsvDate += StartDate.ToString("yyyyMM") + "_list";
                        else if (DateD == "DD") CsvDate += StartDate.ToString("yyyyMMdd") + "_list";

                        //CsvDate = StartDateString;
                    }
                    else
                    {
                        if (DateD == "MM") CsvDate += StartDate.ToString("yyMM") + "到" + EndDate.ToString("yyMM") + "MM_list";
                        else if (DateD == "All") CsvDate += StartDate.ToString("yyMM") + "到" + EndDate.ToString("yyMM") + "_list";
                        else if (DateD == "DD") CsvDate += StartDate.ToString("yyMMdd") + "到" + EndDate.ToString("yyMMdd") + "_list";
                        //CsvDate = StartDate.ToString("yyMMdd") + "到" + EndDate.ToString("yyMMdd");
                    }
                    string FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + "_" + CsvDate + ".csv";
                    //*存文字
                    System.IO.File.WriteAllText(FilePath, lines.ToString(), Encoding.GetEncoding("UTF-8"));
                    FilePathHtml += FilePath;
                }
                else
                {
                    return "無資料匯出";
                }
                return FilePathHtml;

            }
            catch (Exception ex)
            {
                MSSQL.SaveLog(ex.ToString());
                return "錯誤";
            }

        }

        public static string SendMaxMin(SqlCommand cmd, DateTime StartDate, DateTime EndDate, string Where, string WhereJoin, string FAB, string DateD)
        {
            string html = "";
            string whereCountMin = "";
            string whereCountMax = "";
            if (DateD == "MM" || DateD == "DD")
            {
                whereCountMin = @"where doc like @Today and IsFinished = '1' and AliveTime between @TodayF  and  @TodayL and FAB=@FAB";
                whereCountMax = @"where doc like @Today and AliveTime between @TodayF  and  @TodayL and FAB=@FAB";
            }
            else if (DateD == "All")
            {
                whereCountMin = @"where IsFinished = '1' and AliveTime between @TodayF  and @TodayL and FAB=@FAB";
                whereCountMax = @"where AliveTime between  @TodayF  and @TodayL and FAB=@FAB";
                Where = "";
            }
            cmd.CommandText = String.Format(@"select count(doc) from datas {0} {1}", whereCountMin, Where);
            int FinishCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            //總共幾筆
            cmd.CommandText = String.Format(@"select count(doc) from datas {0} {1}", whereCountMax, Where);
            int TotalCount = int.Parse(MSSQL.GetSQLScalar(cmd, Sqlconn));
            if (TotalCount == 0)
            {
                return FAB + "無此資料";
            }
            if (FinishCount == 0) html = FAB + "應完成件數" + TotalCount + "<br/>";
            else html = FAB + "應完成件數" + TotalCount + "<br/>" + "已完成件數" + FinishCount + "<br/>" + "尚未全部完成" + (TotalCount - FinishCount) + "<br/>";

            return html;
        }




        /// <summary>
        /// 報表主要傳值
        /// SendMAll>SendOneList
        /// SendMAll>SendOne
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <param name="DateD"></param>
        /// <param name="TableIDIN"></param>
        /// <param name="TableNameIN"></param>
        /// <returns></returns>
        public static string SendMAll(string FAB, string TxtDate, string DateD, string TableIDIN, string TableNameIN)
        {
            DateTime today = DateTime.Now;

            string Result = "";
            try
            {
                string Today = "", TodayF = "", TodayL = "";
                string where = "";
                string whereJoin = "";
                string whereJoinList = "";

                //當天
                if (DateD == "DD")
                {
                    if (TxtDate != null && TxtDate != "") today = DateTime.Parse(TxtDate);
                    //假日不出單        
                    if (DateSetting.HolidayTest(today) != "") return "今天是假日";
                    string Weekly = today.DayOfWeek.ToString("d");//星期2
                    if (Weekly == "6" || Weekly == "0") return "今天是星期" + Weekly;
                    Today = "%" + today.ToString("yyyyMMdd");
                    TodayF = today.ToString("yyyyMMdd") + " 00:00:00";
                    TodayL = today.ToString("yyyyMMdd") + " 23:59:59";
                }
                else if (DateD == "MM" || DateD == "All")
                {

                    if (TxtDate != null && TxtDate != "")
                    {
                        today = new DateTime(today.Year, int.Parse(TxtDate), 1);
                        TodayF = new DateTime(today.Year, int.Parse(TxtDate), 1).ToString("yyyyMMdd") + " 00:00:00";// 
                        TodayL = new DateTime(today.Year, int.Parse(TxtDate), DateTime.DaysInMonth(today.Year, today.Month)).ToString("yyyyMMdd") + " 23:59:59";//                     
                    }
                    else
                    {
                        today = new DateTime(today.Year, today.Month, 1);
                        TodayF = new DateTime(today.Year, today.Month, 1).ToString("yyyyMMdd") + " 00:00:00";// new DateTime(today.Year, today.Month, 1).ToString("yyyyMMdd") + " 00:00:00";// 
                        TodayL = new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month)).ToString("yyyyMMdd") + " 23:59:59";// new DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month)).ToString("yyyyMMdd") + " 23:59:59";//                     
                    }
                    if (DateD == "MM")
                    {

                        where = " and TableID IN(" + TableIDIN + ")";
                        whereJoin = " and b.TableID IN(" + TableIDIN + ")";
                        whereJoinList = " and c.TableID IN(" + TableIDIN + ")";
                        Today = "%" + today.ToString("yyyyMM") + "M"; //today.AddMonths(-1).ToString("yyyyMMdd");
                    }

                }

                SqlCommand cmd = new SqlCommand();
                cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = Today;
                cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = TodayF;
                cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = TodayL;
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;

                Result = SendMaxMin(cmd, today, today, where, whereJoin, FAB, DateD);
                if (Result == FAB + "無此資料")
                {
                    return Result;
                }
                Result += SendOne(cmd, today, today, where, whereJoin, FAB, DateD);
                Result += SendOneList(cmd, today, today, where, whereJoinList, FAB, DateD);
                return Result.ToString();
            }
            catch
            {
                return "格式不對";
            }
        }
        #endregion
    }
}