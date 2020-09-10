using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    public class DataGetController : Controller
    {

        IDataProvider _IDataProvider;
        public DataGetController()
        {
            _IDataProvider = new DataProvider();
        }

        [HttpGet]
        public ActionResult BarChart()
        {
           
            try
            {
                string tempEnvironmentalConcernCount = string.Empty;
                string tempEnvironmentalConcern = string.Empty;
                //ViewBag.Dashy = _IDataProvider.GetDashboard()
                _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_EnvironmentalConcern.ToString(),DateTime.Now.Year,0,out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);
                ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
                ViewBag.ECName = tempEnvironmentalConcern.Trim();
                //  var x = _IDataProvider.GetDashboard();



                _IDataProvider.GetDashboard();
                ViewBag.lols = _IDataProvider.GetDashboard();
                ViewBag.Greetings = "Hello World!";

              

                return View();
            }
            catch (Exception)
            {
                throw;
            }
        }

        // GET: DataGet
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult GetUserType()
        {
            //= cmd.GetUserType();
            var data  = _IDataProvider.GetUserType();
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetCaseReport(int UpdatedStatusID)
        {
            var data = _IDataProvider.GetCaseReport(UpdatedStatusID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetCurrentCaseReport(int UpdatedStatusID)
        {
            var data = _IDataProvider.GetCaseReport(UpdatedStatusID).Where(x=>x.UpdatedStatusDate.Year == DateTime.Now.Year && x.UpdatedStatusID == UpdatedStatusID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetCurrentCompletedReports(int UpdatedStatusID)
        {
            var data = _IDataProvider.GetCaseReport(UpdatedStatusID).Where(x => x.UpdatedStatusDate.Year == DateTime.Now.Year && x.UpdatedStatusID == UpdatedStatusID);
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetUserInformation()
        {
            var data = _IDataProvider.GetUserInformation().Where(x=>x.UserTypeID==2);
            return Json(data, JsonRequestBehavior.AllowGet);
        }
        public ActionResult GetYear()
        {
            var data = _IDataProvider.GetYearDDL();
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetVolunteer()
        {
            var data = _IDataProvider.GetVolunteer();
            return Json(data, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetUpdatedStatus()
        {
            var data = _IDataProvider.GetUpdatedStatus();
            return Json(data,JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetLeaderboard_Year(int UpdatedStatusID, int Year)
        {
            var data = _IDataProvider.GetLeaderboards_Year(UpdatedStatusID, Year);
            return Json(data, JsonRequestBehavior.AllowGet);
        }
    }
}