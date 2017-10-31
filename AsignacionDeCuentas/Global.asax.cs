using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using BusinessLogic;
using Entity;

namespace AsignacionDeCuentas
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            RouteConfig.RegisterRoutes(RouteTable.Routes);
        }

        protected void Session_Start(object sender, EventArgs e)
        {
            string name = HttpContext.Current.User.Identity.Name;

            if (name.Length > 4)
            {
                name = name.Substring(5);
            }

            Entity.User user = new UserBusiness().GetUserCode(name);
            var session = HttpContext.Current.Session;
            
            if (!string.IsNullOrEmpty(name))
            {
                session["UserCode"] = user.UserCode;
                session["UserName"] = name;
            }
            else
            {
                throw new UnauthorizedAccessException("User not allowed. Bad credential. :(");
            }
            

        }
    }
}
