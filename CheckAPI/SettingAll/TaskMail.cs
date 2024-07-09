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
    public class TaskAddMail
    {
        private static readonly string Sqlconn = ConfigurationManager.AppSettings["SqlConnection"];
       
    }
}