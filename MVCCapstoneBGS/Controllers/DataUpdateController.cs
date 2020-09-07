using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    public class DataUpdateController : Controller
    {

        IDataProvider _IDataProvider;
        public DataUpdateController()
        {
            _IDataProvider = new DataProvider();
        }

        // GET: DataUpdate
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult UpdateVolunteer(int VolunteerID, string GivenName, string MaidenName, string FamilyName)
        {
            var data = _IDataProvider.UpdateVolunteer(VolunteerID, GivenName,MaidenName,FamilyName);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateCaseReportToAccepted(int CaseReportID)
        {
            var data = _IDataProvider.UpdateCaseReport_Accepted(CaseReportID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateCaseReportToRejected(int CaseReportID)
        {
            var data = _IDataProvider.UpdateCaseReport_Rejected(CaseReportID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateCaseReportToCompleted(int CaseReportID)
        {
            var data = _IDataProvider.UpdateCaseReport_Completed(CaseReportID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult UpdateCaseReportToInProgress(int CaseReportID, int VolunteerID)
        {
            var data = _IDataProvider.UpdateCaseReport_InProgress(CaseReportID,VolunteerID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

    }
}