using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    [HandleError]
    public class EntitiesController : Controller
    {
        IDataProvider _IDataProvider;
        public EntitiesController()
        {
            _IDataProvider = new DataProvider();
        }

        string Layout_ADashboard= "~/TerraTech/TerraShared/AdministratorDashboard.cshtml";

        string Layout_CU = "~/TerraTech/TerraShared/CommunityUser.cshtml";
        string Layout_CUDashboard = "~/TerraTech/TerraShared/CommunityUser.cshtml";

        public ActionResult Administrator()
        {
            ViewBag.Title = LabelStruct.Administrator.AdministratorHomepage;
            ViewBag.VBLayout = Layout_ADashboard;
            var LandNumber = _IDataProvider.GetHomeDashboard(DateTime.Now.Year, 1);
            var WaterNumber = _IDataProvider.GetHomeDashboard(DateTime.Now.Year, 2);
            var Users = _IDataProvider.GetDashboard();
            var Progress = _IDataProvider.GetHomeDashboardProgress(DateTime.Now.Year);



            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;          

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_EnvironmentalConcern.ToString(),DateTime.Now.Year,0,out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();

            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_OverallStatusUpdate.ToString(), DateTime.Now.Year, 0, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();

            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

            ViewBag.VBLand = LandNumber;
            ViewBag.VBWater = WaterNumber;
            ViewBag.VBUsers = Users;
            ViewBag.VBProgress = Progress;

            var ggka = _IDataProvider.GetCaseReport(5).ToList();

            string x = string.Empty;
            foreach (var m in ggka)
            {
                x += "[" + m.XCoordinates + ' ' + m.YCoordinates + "]";
               
            }

            ViewBag.DUMMY = x;

            var commaSeparated = string.Join(",", _IDataProvider.GetCaseReport(5).Select(mmm => "["+ mmm.XCoordinates+","+mmm.YCoordinates+"]"));
            ViewBag.DUMMY2 = commaSeparated;


            return View();
        }

        #region CombinedFunctionalities

        public ActionResult Leaderboard()
        {
            ViewBag.Title = LabelStruct.Others.Leaderboards;
            ViewBag.VBLayout = Layout_ADashboard;
            return View();
        }
        #endregion

        #region Administrator
        public ActionResult Accounts()
        {
            ViewBag.Title = LabelStruct.Administrator.Volunteers;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

       
            return View();
        }

        public ActionResult MonthlyReports(int Year = 0, int Month = 0)
        {
            ViewBag.Title = LabelStruct.Administrator.MonthlyReports;
            ViewBag.VBLayout = Layout_ADashboard;

            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_EnvironmentalConcern_MonthYear.ToString(), Year, Month, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();

            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_OverallStatusUpdate_MonthYear.ToString(), Year, Month, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();

            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;
            ViewBag.CURRENT_YEAR = Year;
            ViewBag.CURRENT_MONTH = Month;

            return View();
        }


        public ActionResult YearlyReports(int Year=0, int Month =0)
        {
            ViewBag.Title = LabelStruct.Administrator.YearlyReports;
            ViewBag.VBLayout = Layout_ADashboard;

            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_EnvironmentalConcern.ToString(), Year, Month, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();

            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_OverallStatusUpdate.ToString(), Year, Month, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();

            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;
            ViewBag.CURRENT_YEAR = Year;
            return View();
        }


   

        public ActionResult Twitter()
        {
            ViewBag.Title = LabelStruct.Administrator.Twitter;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

            var LandNumber = _IDataProvider.GetHomeDashboard(DateTime.Now.Year, 1);
            var WaterNumber = _IDataProvider.GetHomeDashboard(DateTime.Now.Year, 2);
            var Users = _IDataProvider.GetDashboard();
            var Progress = _IDataProvider.GetHomeDashboardProgress(DateTime.Now.Year);


            String areaReport="";
            foreach (var dataArea in _IDataProvider.GetAreaDetailsPerMonthYear(DateTime.Now.Month, DateTime.Now.Year))
            {
                areaReport += dataArea.CaseLocation + " = Land: " + dataArea.L_Completed + " Water: " + dataArea.W_Completed + "\n";
            }

            ViewBag.AreaTweet = "Completed Reports per Area: \n"+areaReport;
            
            
            string sentence = string.Empty;

            if (LandNumber > 2)
            {
                sentence = "are";
            }

            ViewBag.UpdateStatus =
                "As of " + DateTime.Now.ToString() + "\nStatus Report:\n" +
                "Completed Land Concern: " + LandNumber + "\n" +
                "Completed Water Concern: " + WaterNumber + "\n" +
                "Pending Cases: " + Progress + "\n" +
                "all of which are taken from the submitted data of " + Users + " users";
            return View();
        }

        public ActionResult Submitted()
        {
            ViewBag.Title = LabelStruct.Administrator.Submitted;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status.ToString(), DateTime.Now.Year, 1, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();


            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status_LW.ToString(), DateTime.Now.Year, 1, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();


            return View();
        }

        public ActionResult Accepted()
        {
            ViewBag.Title = LabelStruct.Administrator.Accepted;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;


            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status.ToString(), DateTime.Now.Year, 3, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();


            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status_LW.ToString(), DateTime.Now.Year, 3, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();



            return View();
        }

        public ActionResult Rejected()
        {
            ViewBag.Title = LabelStruct.Administrator.Rejected;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status.ToString(), DateTime.Now.Year, 2, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();


            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status_LW.ToString(), DateTime.Now.Year, 2, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();

            return View();
        }

        public ActionResult InProgress()
        {
            ViewBag.Title = LabelStruct.Administrator.InProgress;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status.ToString(), DateTime.Now.Year, 4, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();


            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status_LW.ToString(), DateTime.Now.Year, 4, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();

            return View();
        }

        public ActionResult Completed()
        {
            ViewBag.Title = LabelStruct.Administrator.Completed;
            ViewBag.VBLayout = Layout_ADashboard;
            ViewBag.DATETIMENOW = DateTime.Now.Date.ToLongDateString() + " - " + DateTime.Now.TimeOfDay;

            string tempEnvironmentalConcernCount = string.Empty;
            string tempEnvironmentalConcern = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status.ToString(), DateTime.Now.Year, 5, out tempEnvironmentalConcernCount, out tempEnvironmentalConcern);

            ViewBag.ECCount = tempEnvironmentalConcernCount.Trim();
            ViewBag.ECName = tempEnvironmentalConcern.Trim();

            string tempStatusCount = string.Empty;
            string tempStatus = string.Empty;

            _IDataProvider.CHART_Display(StoredProcedureEnum.CHART_Status_LW.ToString(), DateTime.Now.Year, 5, out tempStatusCount, out tempStatus);

            ViewBag.SCount = tempStatusCount.Trim();
            ViewBag.SStatus = tempStatus.Trim();

            return View();
        }
        #endregion

        #region CommunityUser
        public ActionResult CommunityUser(int UpdatedStatusID=0)
        {
            //ViewBag.[Pangalan na gusto mo] = Value na gusto mo;

            ViewBag.Title = LabelStruct.CommunityUser.CommunityUserHomepage;

            var LandNumber = _IDataProvider.GetHomeDashboard(DateTime.Now.Year, 1);
            var WaterNumber = _IDataProvider.GetHomeDashboard(DateTime.Now.Year, 2);
            var Users = _IDataProvider.GetDashboard();
            var Progress = _IDataProvider.GetHomeDashboardProgress(DateTime.Now.Year);


            ViewBag.VBLand = LandNumber.ToString();
            ViewBag.VBWater = WaterNumber;
            ViewBag.VBUsers = Users;
            ViewBag.VBProgress = Progress;


            const string quote = "\"";
            //var commaSeparated = string.Join(",", _IDataProvider.GetCaseReport(5).Select(mmm => "[" + mmm.XCoordinates + "," + mmm.YCoordinates + "]"));

            var commaSeparated = string.Join(",", _IDataProvider.GetCaseReport(UpdatedStatusID).
                Select(
                mmm => "["
                +quote
             //   + "<center><img src='https://i.ytimg.com/vi/EXtNpsj1-0w/hqdefault.jpg' style='width:150px; height:100px;'></center>"
                +"Case No: "+mmm.CaseReportID
                +"<br />Reported on: "+mmm.DateReported
                +"<br />Updated on: "+mmm.UpdatedStatusDate
                +"<br />Type: "+mmm.Concern
                +"<br />City: "+mmm.CaseLocation+" ["+mmm.XCoordinates+","+mmm.YCoordinates+"]"
                +quote
                +","+mmm.XCoordinates+","+mmm.YCoordinates+"]"
                ));
            ViewBag.DUMMY2 = commaSeparated;




            ViewBag.VBLayout = Layout_CUDashboard;
            return View();
        }

        public ActionResult SubmitReport()
        {
            ViewBag.VBLayout = Layout_CU;
            ViewBag.Title = LabelStruct.CommunityUser.SubmitReport;
            return View();
        }

        public ActionResult ViewStatus()
        {
            ViewBag.VBLayout = Layout_CU;
            ViewBag.Title = LabelStruct.CommunityUser.ViewStatus;
            return View();
        }
        public ActionResult Achievements()
        {
            ViewBag.VBLayout = Layout_CU;
            ViewBag.Title = LabelStruct.CommunityUser.Achievements;
            return View();
        }
        #endregion

        #region OtherFunctionalities
        public ActionResult ForgotPassword()
        {

            return View();
        }

        public ActionResult DummyPage1()
        {
            return View();
        }

        public ActionResult DummyPage2()
        {
            return View();
        }

        public ActionResult DummyPage3()
        {
            return View();
        }
        #endregion

    }
}