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
            // Register the BusinessDocuments API
            WebAPIHelper.RegisterWebAPIService(this.HttpContext, "/api/BusinessDocuments");

            return View();
        }
    }
}
