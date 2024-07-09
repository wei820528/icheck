using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace CheckAPI.Models
{
    //public class ExcelMode
    //{
    //    public string Doc { get; set; }
    //    public string FAB { get; set; }
    //    public string TableID { get; set; }
    //    public string AliveTime { get; set; }
    //    public string IsFinished { get; set; }
    //    public string IsFinishedTime { get; set; }
    //}
    public class ExcelDataManage
    {
       
        public string FAB { get; set; }
        public string IsFinished { get; set; }
        public string IsFinishedTime { get; set; }
        public string DateStart { get; set; }
        public string DateEnd { get; set; }
        public string Error { get; set; }
        public string DataName { get; set; }
        
    }
}