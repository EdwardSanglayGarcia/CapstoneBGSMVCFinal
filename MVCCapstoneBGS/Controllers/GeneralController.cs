using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    public class GeneralController : Controller
    {
        IDataProvider _IDataProvider;
        public GeneralController()
        {
            _IDataProvider = new DataProvider();
        }
        // GET: General
        public ActionResult Index()
        {
            return View();
        }
   

        public ActionResult Dummy()
        {
            return View();
        }

        public ActionResult DummyInsert()
        {
            return View();
        }
    }
}