using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CheckAPI.Models
{
    public class AdminLogin_filter : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (filterContext.HttpContext.Session["AdminAccount"] == null || filterContext.HttpContext.Session["AdminAccount"].ToString() == "")
            {
                UrlHelper uu = new UrlHelper(filterContext.Controller.ControllerContext.RequestContext);
                string url = uu.Action("LoginIndex", "Login");
                filterContext.HttpContext.Response.Redirect(url);
            }
        }
    }
}