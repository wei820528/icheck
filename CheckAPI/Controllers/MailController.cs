using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace CheckAPI.Controllers
{
    public class MailController : Controller
    {
        // GET: Mail
        public static string SendMailByGmail(List<string> MailList, string Subject, string Body)
        {
            string MailAccount = "", MailAccountPwd = "", Smtp = "", SendMailFrom = "", SendMailTitle = "";
            int SmtpPort = 0; bool EnableSsl = true;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"select * from Mail_Setting where MailNo='01'");
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            if (dt.Rows.Count < 1) return " Mail設定錯誤";
            foreach (DataRow dr in dt.Rows)
            {
                MailAccount = dr["MailAccount"].ToString();
                MailAccountPwd = dr["MailAccountPwd"].ToString();
                Smtp = dr["MailSmtp"].ToString();
                SmtpPort = int.Parse(dr["MailSmtpPort"].ToString());
                EnableSsl = bool.Parse(dr["EnableSsl"].ToString());
                SendMailFrom = dr["SendMailFrom"].ToString();
                SendMailTitle = dr["SendMailTitle"].ToString();
                break;
            }
            try
            {
                MailMessage msg = new MailMessage();
                //收件者，以逗號分隔不同收件者 ex "test@gmail.com,test2@gmail.com"
                msg.To.Add(string.Join(",", MailList.ToArray()));
                msg.From = new MailAddress(SendMailFrom, SendMailTitle, Encoding.UTF8);
                //郵件標題 
                msg.Subject = Subject;
                //郵件標題編碼  
                msg.SubjectEncoding = Encoding.UTF8;
                //郵件內容
                msg.Body = Body;
                msg.IsBodyHtml = true;
                msg.BodyEncoding = Encoding.UTF8;//郵件內容編碼 
                msg.Priority = MailPriority.Normal;//郵件優先級 
                                                   //建立 SmtpClient 物件 並設定 Gmail的smtp主機及Port 


                #region 其它 Host
                /*
                 *  outlook.com smtp.live.com port:25
                 *  yahoo smtp.mail.yahoo.com.tw port:465
                */
                #endregion
                SmtpClient MySmtp = new SmtpClient(Smtp, SmtpPort);
                //SmtpClient MySmtp = new SmtpClient("smtp.gmail.com", 25);
                //設定你的帳號密碼
                MySmtp.Credentials = new System.Net.NetworkCredential(MailAccount, MailAccountPwd);
                //Gmial 的 smtp 使用 SSL
                MySmtp.EnableSsl = EnableSsl;
                MySmtp.Send(msg);
                return "";
            }
            catch (Exception ex)
            {
                string e = ex.ToString();
                SaveLog(e);
                return " Mail設定錯誤";
            }
        }
        // GET: Mail
        public static string SendMailGmailFiles(List<string> MailList, string Subject, string Body, List<string> FileUrlArray)
        {
            string MailAccount = "", MailAccountPwd = "", Smtp = "", SendMailFrom = "", SendMailTitle = "";
            int SmtpPort = 0; bool EnableSsl = true;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"select * from Mail_Setting where MailNo='01'");
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            if (dt.Rows.Count < 1) return " Mail設定錯誤";
            foreach (DataRow dr in dt.Rows)
            {
                MailAccount = dr["MailAccount"].ToString();
                MailAccountPwd = dr["MailAccountPwd"].ToString();
                Smtp = dr["MailSmtp"].ToString();
                SmtpPort = int.Parse(dr["MailSmtpPort"].ToString());
                EnableSsl = bool.Parse(dr["EnableSsl"].ToString());
                SendMailFrom = dr["SendMailFrom"].ToString();
                SendMailTitle = dr["SendMailTitle"].ToString();
                break;
            }
            try
            {
                MailMessage msg = new MailMessage();
                //收件者，以逗號分隔不同收件者 ex "test@gmail.com,test2@gmail.com"
                msg.To.Add(string.Join(",", MailList.ToArray()));
                msg.From = new MailAddress(SendMailFrom, SendMailTitle, Encoding.UTF8);
                //郵件標題 
                msg.Subject = Subject;
                //郵件標題編碼  
                msg.SubjectEncoding = Encoding.UTF8;
                //郵件內容
                msg.Body = Body;
                msg.IsBodyHtml = true;
                msg.BodyEncoding = Encoding.UTF8;//郵件內容編碼 
                msg.Priority = MailPriority.Normal;//郵件優先級 
                                                   //建立 SmtpClient 物件 並設定 Gmail的smtp主機及Port 

                //附加檔案
                if (FileUrlArray != null && FileUrlArray.Count > 0)
                {
                    foreach (string file in FileUrlArray)
                    {
                        msg.Attachments.Add(new Attachment(file)); //加入附加檔案
                    }
                }


                #region 其它 Host
                /*
                 *  outlook.com smtp.live.com port:25
                 *  yahoo smtp.mail.yahoo.com.tw port:465
                */
                #endregion
                SmtpClient MySmtp = new SmtpClient(Smtp, SmtpPort);
                //SmtpClient MySmtp = new SmtpClient("smtp.gmail.com", 25);
                //設定你的帳號密碼
                MySmtp.Credentials = new System.Net.NetworkCredential(MailAccount, MailAccountPwd);
                //Gmial 的 smtp 使用 SSL
                MySmtp.EnableSsl = EnableSsl;
                MySmtp.Send(msg);
                return "";
            }
            catch (Exception ex)
            {
                string e = ex.ToString();
                SaveLog(e);
                return " Mail設定錯誤";
            }
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
        public static Encoding GetByte(byte[] filebyte, MemoryStream blobStream)
        {
            System.IO.BinaryReader br = new System.IO.BinaryReader(blobStream);
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
        private static string BytesToStringConverted(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes))
            {
                using (var streamReader = new StreamReader(stream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }
        public static string SendMailGmailFilesTest(List<string> MailList, string Subject, string Body,string Txt,string TxtName)
        {
            string MailAccount = "", MailAccountPwd = "", Smtp = "", SendMailFrom = "", SendMailTitle = "";
            int SmtpPort = 0; bool EnableSsl = true;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = string.Format(@"select * from Mail_Setting where MailNo='01'");
            DataTable dt = GetSQLDataTable(cmd, Sqlconn);
            if (dt.Rows.Count < 1) return " Mail設定錯誤";
            foreach (DataRow dr in dt.Rows)
            {
                MailAccount = dr["MailAccount"].ToString();
                MailAccountPwd = dr["MailAccountPwd"].ToString();
                Smtp = dr["MailSmtp"].ToString();
                SmtpPort = int.Parse(dr["MailSmtpPort"].ToString());
                EnableSsl = bool.Parse(dr["EnableSsl"].ToString());
                SendMailFrom = dr["SendMailFrom"].ToString();
                SendMailTitle = dr["SendMailTitle"].ToString();
                break;
            }
            try
            {

                byte[] blobData = Encoding.UTF8.GetBytes(Txt);
                MemoryStream blobStream = new MemoryStream(blobData); // 創建Blob流

                //var encoding = GetByte(blobData, blobStream);
                //string a = BytesToStringConverted(blobData);
                //string result = System.Text.Encoding.UTF8.GetString(blobData);
                  
                
                //Encoding.GetEncoding("UTF-8").GetBytes(Txt); //Encoding.UTF8.GetBytes(Txt); // 將純文字轉換為Blob數據
               
                
                MailMessage msg = new MailMessage();
                //收件者，以逗號分隔不同收件者 ex "test@gmail.com,test2@gmail.com"
                msg.To.Add(string.Join(",", MailList.ToArray()));
                msg.From = new MailAddress(SendMailFrom, SendMailTitle, Encoding.UTF8);
                //郵件標題 
                msg.Subject = Subject;
                //郵件標題編碼  
                msg.SubjectEncoding = Encoding.UTF8;
                //郵件內容
                msg.Body = Body;
                msg.IsBodyHtml = true;
                msg.BodyEncoding = Encoding.UTF8;//郵件內容編碼 
                msg.Priority = MailPriority.Normal;//郵件優先級 
                                                   //建立 SmtpClient 物件 並設定 Gmail的smtp主機及Port 
                                                   // 添加Blob附件

             //   Attachment attachment = new Attachment(blobStream, TxtName+".csv", "text/vnd.ms-excel");
                Attachment attachment = Attachment.CreateAttachmentFromString(
                                 Txt,
                                TxtName + ".csv",
                                Encoding.Default,
                               "text/vnd.ms-excel");
                msg.Attachments.Add(attachment);

                #region 其它 Host
                /*
                 *  outlook.com smtp.live.com port:25
                 *  yahoo smtp.mail.yahoo.com.tw port:465
                */
                #endregion
                SmtpClient MySmtp = new SmtpClient(Smtp, SmtpPort);
                //SmtpClient MySmtp = new SmtpClient("smtp.gmail.com", 25);
                //設定你的帳號密碼
                MySmtp.Credentials = new System.Net.NetworkCredential(SendMailFrom, MailAccountPwd);
                //MySmtp.UseDefaultCredentials = false;
                //Gmial 的 smtp 使用 SSL
                MySmtp.EnableSsl = EnableSsl;
                MySmtp.Send(msg);


                // 释放资源
                attachment.Dispose();
                blobStream.Dispose();



                return "";
            }
            catch (Exception ex)
            {
                string e = ex.ToString();
                SaveLog(e);
                return " Mail設定錯誤";
            }

        }
        //public class MailModel
        //{
        //    public string Sender { get; set; } //寄件者
        //    public string SenderName { get; set; } //顯示名稱
        //    public string Recipient { get; set; } //接收者
        //    public string RecipientName { get; set; } //顯示名稱
        //    public string Title { get; set; } //信件標題
        //    public string Content { get; set; } //信件內容
        //    public string SMTP_Host { get; set; }
        //    public int SMTP_Port { get; set; }
        //    public List<string> Files { get; set; }
        //}

        //public void SendMail(MailModel Model)
        //{
        //    MailAddress from = new MailAddress(Model.Sender, Model.SenderName, Encoding.UTF8); //地址, 顯示名稱, 編碼方式
        //    MailAddress to = new MailAddress(Model.Recipient, Model.RecipientName, Encoding.UTF8); //地址, 顯示名稱
        //    MailMessage message = new MailMessage(from, to);  //MailMessage(寄信者, 收信者)

        //    //message.From = new MailAddress(Model.Sender, Model.SenderName, Encoding.UTF8);

        //    message.BodyEncoding = Encoding.UTF8; //E-mail編碼
        //    message.Subject = Model.Title; //E-mail主旨
        //    message.SubjectEncoding = Encoding.UTF8; //設定信箱主旨編碼方式
        //    message.Body = Model.Content; //E-mail內容
        //    message.BodyEncoding = Encoding.UTF8; //設定信箱內容編碼方式
        //    message.IsBodyHtml = true; //啟用 HTML 格式
        //    message.Priority = MailPriority.Normal; //設定優先權 MailPriority.Hight 會設為重要信件
        //                                            //message.Headers.Add("Return-Path", "*****@*****.*****"); //退信位置 

        //    //附加檔案
        //    if (Model.Files != null && Model.Files.Count > 0)
        //    {
        //        foreach (string file in Model.Files)
        //        {
        //            message.Attachments.Add(new Attachment(file)); //加入附加檔案
        //        }
        //    }

        //    SmtpClient smtpClient = new SmtpClient(Model.SMTP_Host, Model.SMTP_Port);    //設定E-mail Server和port
        //    smtpClient.EnableSsl = false; // 是否啟用 SSL，Gmial需要啟用 SSL
        //    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network; // 指定外送電子信箱的處理方式
        //                                                            //smtpClient.Credentials = new System.Net.NetworkCredential("Account","Password"); 
        //                                                            //利用帳號／密碼取得 Smtp 伺服器的憑證 

        //    try
        //    {
        //        smtpClient.Send(message);// 發送
        //        smtpClient.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        //無法使用信箱。 伺服器回應為: User (*****@*****.*****) unknown.
        //        //Console.WriteLine("Exception caught in XXXX: {0}",ex.ToString());
        //        //錯誤處理和紀錄
        //    }
        //    message.Dispose();//放掉宣告出來的MailMessage

        //}


        #region SQLCmd
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
        private static string GetSQLScalar(SqlCommand cmd, string Sqlconn)
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
        private static string GetSQLNonQuery(SqlCommand cmd, string Sqlconn)
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
                Result = "";
            }
            return Result;
        }
        private static DataTable GetSQLDataTable(SqlCommand cmd, string Sqlconn)
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
            catch { }
            return dt;
        }
        private static void SaveLog(string json)
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
        #endregion
    }
}