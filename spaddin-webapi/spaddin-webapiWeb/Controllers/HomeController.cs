using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.WebAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace spaddin_webapiWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            // Register the Web API when accessing the default page
            Register();
            
            return View();
        }

        [SharePointContextFilter]
        public ActionResult Register()
        {
            try
            {
                // Register the BusinessDocuments API
                WebAPIHelper.RegisterWebAPIService(this.HttpContext, "/api/BusinessDocuments");
                return Json(new { message = "Web API registered" });
            }
            catch (Exception ex)
            {
                return Json(ex);
            }
        }
    }
}
