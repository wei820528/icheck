using CheckAPI.Controllers;
using CheckAPI.Models;
using Microsoft.SqlServer.Server;
using System;
using System.Collections;
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
    public class TextCsv
    {
        private static string JoinHtmlThead(string[] HtmlArray) {
            string htmlD = "";
            for (int i = 0; i < HtmlArray.Length; i++)
            {
                htmlD += @"<th scope=""col"">"+ HtmlArray[i] + "</th>";
            }
            return htmlD;
        }
        #region 共用主
        /// <summary>
        /// 日/月 csv
        /// </summary>
        /// <param name="FAB">廠區</param>
        /// <param name="TxtDate">開始日期</param>
        /// <param name="returnString">製作的文字</param>
        /// <param name="FormatData">""/"M"/"Y"/"all"</param>
        /// <param name="Complete">"0"未完成/"1"完成/""全部</param>
        /// <returns></returns>
        private static string Save_Csv_Form(string FAB, string TxtDate, string returnString, string FormatData, string Complete)
        {
            DateTime today = SendYMDController.DateString(TxtDate);
            StringBuilder lines = new StringBuilder();
            StringBuilder linelists = new StringBuilder();
            //製作，已完成未完成總量數據
                       
            //普通訊息
            string line = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""巡檢人員"",=""人員工號"",=""規格表單編號""";
            string linehtml = "";
            lines.AppendLine(line);
            //詳細訊息list
            string lineList = @"=""廠別"",=""巡檢單號"",=""巡檢表單名稱"",=""巡檢時間"",=""人員工號"",=""巡檢人員"",=""項目名稱"",=""巡檢紀錄"",=""是否異常"",=""異常原因"",=""項目說明"",=""規格表單編號""";
            string lineListhtml = "";
            linelists.AppendLine(lineList);
            DataTable DBList = new DataTable();
            DataTable DB = new DataTable();
            linehtml += "<table>廠區" + FAB + "表單";
            lineListhtml += "<table>廠區" + FAB + "表單明細";
            string FilePath = "";
            string FilePathList = "";

            FilePath = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + today.ToString("yyyyMMdd") + ".csv";
            FilePathList = ConfigurationManager.AppSettings["TxtUrl"] + FAB + @"_" + today.ToString("yyyyMMdd") + "list.csv";
            if (FormatData == "M" || FormatData == "all" || FormatData == "")
            {
                //日//已完成/未完成
                DBList = Today_Form_all_list(FAB, TxtDate, TxtDate, Complete, true, FormatData);
                int a = DBList.Rows.Count;
                DB = Today_Form_all_list(FAB, TxtDate, TxtDate, Complete, false, FormatData);
                int b = DB.Rows.Count;
            }
            else {
                FilePath = "";
            }

            if (DBList.Rows.Count == 0 || DB.Rows.Count == 0) {
                if (DBList.Rows.Count == 0)
                {
                    linehtml = "<thead><tr>無表單</tr></thead></table>";
                }
                if (DB.Rows.Count == 0)
                {
                    lineListhtml = "<thead><tr>無詳細表單</tr></thead></table>";
                }
                line = "<br/>";
                return returnString + linehtml + "，" + lineListhtml + line;
            }
            /////注意
            //csv製作
            linelists = All_CSV_List_Save(DBList, lineList);
            SaveName(FilePathList, linelists, "UTF-8");
            lines = All_CSV_Save( DB, line);
            SaveName(FilePath, lines, "UTF-8");
            string[] linehtmlArray = new string[] { "廠別", "巡檢單號", "巡檢表單名稱", "巡檢時間", "巡檢人員", "人員工號", "規格表單編號" };
            //html製作
            linehtml += "<thead><tr>";
            linehtml += JoinHtmlThead(linehtmlArray);//表頭
            linehtml += "</tr></thead>";
            linehtml += "<tbody>";
            linehtml = All_Html_Save(DB, linehtml);//表尾
            linehtml += "</tbody>";            
            linehtml += "</table>";

            string[] lineListhtmlArray = new string[] { "廠別", "巡檢單號", "巡檢表單名稱", "巡檢時間", "人員工號", "巡檢人員", "項目名稱", "巡檢紀錄", "是否異常", "異常原因", "項目說明",  "規格表單編號" };
            lineListhtml += "<thead><tr>";
            lineListhtml += JoinHtmlThead(lineListhtmlArray);
            lineListhtml += "</tr></thead>";
            lineListhtml += "<tbody>";
            lineListhtml = All_Html_List_Save(DBList, lineListhtml);//表尾
            lineListhtml += "</tbody>";           
            lineListhtml += "</table>";
           
            line = "<br/>";
            return returnString+linehtml + line + lineListhtml;
        }
        //製作list CSV
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt">詳細資料庫內容</param>
        /// <param name="line">之前存的內容</param>
        /// <returns></returns>
        private static StringBuilder All_CSV_List_Save(DataTable dt,string line)
        {
            StringBuilder lines = new StringBuilder();
            if (dt.Rows.Count != 0)
            {
                //表頭
               
                lines.AppendLine(line);
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
                }
               
            }
            return lines;
        }
        //製作CSV
        private static StringBuilder All_CSV_Save( DataTable dt, string line)
        {
            StringBuilder lines = new StringBuilder();
            if (dt.Rows.Count != 0)
            {
                lines.AppendLine(line);
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
                }

            }
            return lines;
        }
        //製作list Html
        private static string All_Html_List_Save(DataTable dt, string line)
        {
            if (dt.Rows.Count != 0)
            {
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



                    line += @"<tr>";
                    line += @"<td>" + dr["FAB"].ToString() + @"</td>";
                    line += @"<td>" + dr["DOC"].ToString() + @"</td>";
                    line += @"<td>" + dr["TableName"].ToString() + @"</td>";
                    if (dr["IsFinishedTime"].ToString() != "")
                    {
                        line += @"<td>" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @"</td>";

                    }
                    else line += @"<td></td>";
                    line += @"<td>" + dr["UserID"].ToString() + @"</td>";
                    line += @"<td>" + dr["UserName"].ToString() + @"</td>";
                    line += @"<td>" + ItemNameTxt + @"</td>";
                    if (dr["ItemValue"].ToString() == "OK")//巡檢紀錄
                        line += @"<td>正常</td>";
                    else
                        line += @"<td>" + dr["ItemValue"].ToString() + @"</td>";


                    line += @"<td>" + IsError + @"</td>";
                    line += @"<td>" + ErrorMsg + @"</td>";
                    line += @"<td>" + dr["ItemContent"].ToString() + @"</td>";
                    line += @"<td>" + dr["SerialNumber"].ToString() + @"</td>";
                    line += @"</tr>";

                }

            }
            return line;
        }
        //製作Html
        private static string All_Html_Save(DataTable dt, string  line)
        {
            if (dt.Rows.Count != 0)
            {
                //*
                foreach (DataRow dr in dt.Rows)
                {
                    line += @"<tr>";
                    line += @"<td>" + dr["FAB"].ToString() + @"</td>";
                    line += @"<td>" + dr["DOC"].ToString() + @"</td>";
                    line += @"<td>" + dr["TableName"].ToString() + @"</td>";
                    if (dr["IsFinishedTime"].ToString() != "")
                    {
                        line += @"<td>" + DateTime.Parse(dr["IsFinishedTime"].ToString()).ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.CreateSpecificCulture("en-US")) + @"</td>";

                    }
                    else line += @"<td></td>";

                    line += @"<td>" + dr["UserName"].ToString() + @"</td>";
                    line += @"<td>" + dr["UserID"].ToString() + @"</td>";
                    line += @"<td>" + dr["SerialNumber"].ToString() + @"</td>";
                    line += @"</tr>";
                }

            }
            return line;

        }
        #endregion
        #region 共用
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        /// <summary>
        /// 匯處csv
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="lines"></param>
        /// <param name="languageData"></param>
        /// <returns></returns>
        public static string SaveName(string FilePath, StringBuilder lines, string languageData)
        {
            string Result = "ok";
            try
            {
                System.IO.File.WriteAllText(FilePath, lines.ToString(), Encoding.GetEncoding(languageData));
            }
            catch (Exception ex)
            {
                MSSQL.SaveLog(ex.ToString());
            }

            return Result;
        }
        /// <summary>
        /// 報表查詢(日/月)
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="where"></param>
        /// <returns></returns>
        public static DataTable Today_Form_All(SqlCommand cmd, string where)
        {
            DataTable dt = new DataTable();
            cmd.CommandText = String.Format(@"
                    select a.FAB, a.Doc,b.TableName, a.IsFinishedTime, isnull(a.UserID,b.UserID) as UserID, c.UserName,d.SerialNumber,d.TableType
                    from datas a
                    left join tables b on b.TableID=a.TableID
                    left join users c on c.UserID=isnull(a.UserID,b.UserID)
                    left join TablesType as d on b.TableType=d.TableType                   
                    {0}
                    ", where);
            dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }
        /// <summary>
        /// 報表詳細查詢(日/月)
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="where"></param>
        /// <returns></returns>
        public static DataTable Today_Form_All_List(SqlCommand cmd, string where)
        {
            DataTable dt = new DataTable();
            cmd.CommandText = String.Format(@"
                    select a.FAB,a.Doc,c.TableName,a.IsFinishedTime,a.UserID,e.UserName,d.ItemName,b.ItemValue,d.ItemType,d.ItemMax,d.ItemMin,f.SerialNumber,f.TableType,d.ItemContent
                    from datas as a
                    left join DatasItem as b	on a.doc = b.doc
                    left join Tables as c		on a.TableID = c.TableID
                    left join TablesItem as d	on a.TableID=d.TableID and b.ItemID=d.ItemID
                    left join Users as e		on a.UserID=e.UserID 
                    left join TablesType as f		on c.TableType=f.TableType
                    {0}
                    ", where);
            dt = MSSQL.GetSQLDataTable(cmd, Sqlconn);
            return dt;
        }


        /// <summary>
        /// 更新完成數據
        /// </summary>
        /// <param name="FAB">廠區</param>
        /// <param name="Format">""/"M"/"Y"</param>
        /// <param name="Complete">0未完成/1完成</param>
        /// <returns></returns>
        private static string Factories_Complete(string FAB, string Format, string Complete,string TxtDate)
        {
            DateTime today = SendYMDController.DateString(TxtDate);
            string where = "WHERE FAB=@FAB";
            if (FAB =="")  where = "";
            //DateTime today = DateTime.Now;
            string ok = "";
            SqlCommand cmd = new SqlCommand();
            cmd.Parameters.Add("@Complete_Time", SqlDbType.VarChar).Value = today.ToString("yyyy-MM-dd HH:mm:ss");
            cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            cmd.Parameters.Add("@Complete", SqlDbType.Bit).Value = int.Parse(Complete);
            cmd.Parameters.Add("@CreateDate", SqlDbType.Date).Value = today.ToString("yyyy-MM-dd");
            if (Format == "all")
            {
                cmd.CommandText = string.Format(@"
                UPDATE Factories
                SET Complete_Time = @Complete_Time,Complete = @Complete,CreateDate = @CreateDate,
                    CompleteM_Time = @Complete_Time,CompleteM = @Complete,
                    CompleteY_Time = @Complete_Time,CompleteY = @Complete,
                {0}", where);
            }
            else {
                cmd.CommandText = string.Format(@"
                UPDATE Factories
                SET Complete{1}_Time = @Complete_Time,Complete{1} = @Complete,CreateDate = @CreateDate
                {0}", where, Format);
            }

            ok = MSSQL.GetSQLNonQuery(cmd, Sqlconn);
            return ok;
        }
        #endregion
        #region 主要月

        /// <summary>
        /// 產生月CSV
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string All_CSV_M(string FAB, string TxtDate, string returnString)
        {
            //Result += dr["FAB"].ToString() + "<br/>";
            int HH = int.Parse(DateTime.Now.ToString("HH"));
            //全部印
            if (HH >= int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                //時間到全部印!並且改1
               
                returnString += FAB + "已做完表單<br>";
                returnString = Save_Csv_Form(FAB,TxtDate,returnString,"M", "");
                Factories_Complete(FAB, "M", "1", TxtDate);
            }
            else if (HH < int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                //查詢未完成件數
                int count = Count_NoEnd_all_Form(FAB, TxtDate, TxtDate,"M");
                if (count == 0)
                {
                    
                    returnString += FAB + "已做完表單<br>";
                    returnString = Save_Csv_Form(FAB, TxtDate, returnString, "M", "1");
                    Factories_Complete(FAB, "M", "1", TxtDate);
                }
                else
                {
                   
                    returnString += FAB + "未做完表單<br>";
                    Factories_Complete(FAB, "M", "0", TxtDate);
                }
            }

            return returnString;
        }
        #endregion

        #region 主要全部

        #endregion
        //已完成日件數
        public static int Count_End_all_Form(string FAB, string TxtDate, string TxtDateEnd, string FormatData)
        {
            DataTable DBList = Today_Form_all_list(FAB, TxtDate, TxtDateEnd, "1", true, FormatData);
            DataTable DB = Today_Form_all_list(FAB, TxtDate, TxtDateEnd, "1", false, FormatData);
            return DB.Rows.Count;
        }
        //未完成日件數0.
        public static int Count_NoEnd_all_Form(string FAB, string TxtDate, string TxtDateEnd, string FormatData)
        {
            DataTable DBList = Today_Form_all_list(FAB, TxtDate, TxtDateEnd, "0", true, FormatData);
            DataTable DB = Today_Form_all_list(FAB, TxtDate, TxtDateEnd, "0", false, FormatData);
            int a = DBList.Rows.Count;
            int b = DB.Rows.Count;
            return DB.Rows.Count;
        }
        //全部日件數
        public static int Count_All_all_Form(string FAB, string TxtDate, string TxtDateEnd, string FormatData)
        {
            DataTable DBList = Today_Form_all_list(FAB, TxtDate, TxtDateEnd, "", true, FormatData);
            DataTable DB = Today_Form_all_list(FAB, TxtDate, TxtDateEnd, "", false, FormatData);
            return DB.Rows.Count;
        }
        /// <summary>
        /// 全部詳細訊息
        /// </summary>
        /// <param name="FAB">廠區</param>
        /// <param name="TxtDate">一開始日期</param>
        /// <param name="TxtDateEnd">最後日期</param>
        /// <param name="IsFinished">"0"未完成/"1"已完成/""全部</param>
        /// <param name="ListBool">false沒list/true有list</param>
        /// <param name="FormatData">""/"M"/"Y"/"all"</param>
        /// <returns></returns>
        public static DataTable Today_Form_all_list(string FAB, string TxtDate, string TxtDateEnd, string IsFinished, bool ListBool, string FormatData)
        {
            All_List.FAB = FAB;
            All_List.TxtDate = TxtDate;
            DateTime today = SendYMDController.DateString(TxtDate);
            DateTime todayEnd = SendYMDController.DateString(TxtDateEnd);
            SqlCommand cmd = new SqlCommand { };
            string a = today.ToString("yyyyMMdd") + " 00:00:00";
            string b = todayEnd.ToString("yyyyMMdd") + " 23:59:59";
            string whereToday = "";
            if (FormatData == "")
            {
                whereToday = "where a.doc like '%' + @Today ";
                cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
            }
            else if (FormatData == "all")
            {
                whereToday = "where (a.doc like '%' + @Today or a.doc like '%' + @TodayM + 'M') ";
                cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd");
                cmd.Parameters.Add("@TodayM", SqlDbType.VarChar).Value = today.ToString("yyyyMM");
            }
            else if (FormatData == "M")
            {
                whereToday = "where a.doc like '%' + @Today + 'M' ";
                cmd.Parameters.Add("@Today", SqlDbType.VarChar).Value = today.ToString("yyyyMM");
            }
            else if (FormatData == "Y")
            {
            }
            else
            {
                DataTable errorDB = new DataTable();
                return errorDB;
            }



            string whereIsFinished = "";
            if (IsFinished != "")
            {
                int number;
                if (int.TryParse(IsFinished, out number))
                {
                    //if (number == 1)
                    //{
                    //    cmd.Parameters.Add("@IsFinished", SqlDbType.VarChar).Value = IsFinished;
                    //    cmd.Parameters.Add("@TodayF", SqlDbType.VarChar).Value = today.ToString("yyyyMMdd") + " 00:00:00";
                    //    cmd.Parameters.Add("@TodayL", SqlDbType.VarChar).Value = todayEnd.ToString("yyyyMMdd") + " 23:59:59";
                    //    whereIsFinished = "and a.IsFinishedTime >= @TodayF and a.IsFinishedTime <= @TodayL ";
                    //}
                    //else 
                    cmd.Parameters.Add("@IsFinished", SqlDbType.VarChar).Value = IsFinished;
                }
                else
                {
                    cmd.Parameters.Add("@IsFinished", SqlDbType.VarChar).Value = "0";
                }
                whereIsFinished += "and a.IsFinished = @IsFinished";
            }
            string whereFAB = "";
            if (FAB != "")
            {
                whereFAB = " and a.FAB = @FAB ";
                cmd.Parameters.Add("@FAB", SqlDbType.VarChar).Value = FAB;
            }
            string where = String.Format(@" {0} {1} {2}", whereToday, whereIsFinished, whereFAB);
            // return Today_Form_All_List(cmd, where);
            if (ListBool) return Today_Form_All_List(cmd, where);
            else return Today_Form_All(cmd, where);
        }


        /// <summary>
        /// 產生月CSV
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string All_CSV_All (string FAB, string TxtDate, string returnString, string FormatData)
        {
            int HH = int.Parse(DateTime.Now.ToString("HH"));
            bool OKholiday = DateSetting.FinalDay(TxtDate);
           
            if (OKholiday)
            {
                FormatData = "all";
            }
            else {
                FormatData = "";
            }
            //全部印
            if (HH >= int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                //時間到全部印!並且改1               
                returnString += FAB + "已做完表單<br>";
                returnString = Save_Csv_Form(FAB, TxtDate, returnString, FormatData, "");
                Factories_Complete(FAB, "", "1", TxtDate);
            }
            else if (HH < int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                //查詢未完成件數
                int count = Count_NoEnd_all_Form(FAB, TxtDate, TxtDate, FormatData);
                if (count == 0)
                {
                    returnString += FAB + "已做完表單<br>";
                    returnString = Save_Csv_Form(FAB, TxtDate, returnString, FormatData, "1");
                    Factories_Complete(FAB, "", "1", TxtDate);
                }
                else
                {
                    returnString += FAB + "未做完表單<br>";
                    Factories_Complete(FAB, "", "0", TxtDate);
                }
            }

            return returnString;
        }


        #region 主要日



        /// <summary>
        /// 產生月CSV
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        public static string All_CSV_D(string FAB, string TxtDate, string returnString)
        {
            //Result += dr["FAB"].ToString() + "<br/>";
            int HH = int.Parse(DateTime.Now.ToString("HH"));

            //全部印
            if (HH >= int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                //時間到全部印!並且改1               
                returnString += FAB + "已做完表單<br>";
                returnString = Save_Csv_Form(FAB, TxtDate, returnString, "", "");
                Factories_Complete(FAB, "", "1", TxtDate);
            }
            else if (HH < int.Parse(ConfigurationManager.AppSettings["Effective_Time"]))
            {
                //查詢未完成件數
                int count = Count_NoEnd_all_Form(FAB, TxtDate, TxtDate,"");
                if (count == 0)
                {                    
                    returnString += FAB + "已做完表單<br>";
                    returnString = Save_Csv_Form(FAB, TxtDate, returnString, "", "1");
                    Factories_Complete(FAB, "", "1", TxtDate);
                }
                else
                {
                    returnString += FAB + "未做完表單<br>";
                    Factories_Complete(FAB, "", "0", TxtDate);
                }
            }

            return returnString;
        }

        #endregion
        /// <summary>
        /// 產生CSV
        /// </summary>
        /// <param name="FAB"></param>
        /// <param name="TxtDate"></param>
        /// <returns></returns>
        private string All_CSV_Save_1(string FAB, string TxtDate)
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
                select count(doc)  from datas where doc like '%'+@Today and FAB=@FAB"
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
            }
            if (dt.Rows.Count != 0)
            {
                //表頭
                StringBuilder lines = new StringBuilder();
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


    }
}