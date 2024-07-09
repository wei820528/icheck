using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CheckAPI.Models
{
    public class UserLogin_filter : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (filterContext.HttpContext.Session["LoginAccount"] == null || filterContext.HttpContext.Session["LoginAccount"].ToString() == "")
            {
                UrlHelper uu = new UrlHelper(filterContext.Controller.ControllerContext.RequestContext);
                string url = uu.Action("LoginIndex", "Login");
                filterContext.HttpContext.Response.Redirect(url);
            }
        }
    }
}