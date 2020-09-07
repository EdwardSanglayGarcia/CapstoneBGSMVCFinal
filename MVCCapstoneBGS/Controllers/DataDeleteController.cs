using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    public class DataDeleteController : Controller
    {

        IDataProvider _IDataProvider;
        public DataDeleteController()
        {
            _IDataProvider = new DataProvider();
        }

        // GET: DataDelete
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult DeleteUserType(int UserTypeID)
        {
            var data = _IDataProvider.DeleteUserType(UserTypeID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult DeleteVolunteer(int VolunteerID)
        {
            var data = _IDataProvider.DeleteVolunteer(VolunteerID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }
    }
}