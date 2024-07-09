using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CheckAPI.Controllers
{
    public class ExcelDownloadAllController : Controller
    {
        // GET: ExcelDownloadAll
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// (錯誤表單)
        /// </summary>
        /// <returns></returns>
        public string ErrorForm() {

            //巡檢表單錯誤

            return "";
        }
        /// <summary>
        /// (正確表單)
        /// </summary>
        /// <returns></returns>
        public string TrueForm()
        {

            //巡檢表單正確

            return "";
        }
        /// <summary>
        /// (全部表單)
        /// </summary>
        /// <returns></returns>
        public string ALLForm()
        {

            //巡檢表單正確

            return "";
        }
    }
}