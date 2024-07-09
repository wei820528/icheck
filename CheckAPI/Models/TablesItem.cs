using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CheckAPI.Models
{
    public class TablesItem
    {
        public string FAB { get; set; }
        public string TableID { get; set; }
        public string ItemID { get; set; }
        public int ItemSort { get; set; }
        public string ItemName { get; set; }
        public string ItemContent { get; set; }
        public string ItemType { get; set; }
        public int ItemMin { get; set; }
        public int ItemMax { get; set; }
        public DateTime CreateTime { get; set; }
    }
}