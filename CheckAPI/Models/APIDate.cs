using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CheckAPI.Models
{
    public class APIDate
    {
        public DateTime date { get; set; }
        public string chinese { get; set; }
        public bool isholiday { get; set; }
        public string holidaycategory { get; set; }

        public string description { get; set; }
    }
}