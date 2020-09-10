using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCCapstoneBGS
{
    public class UserType
    {
        public int UserTypeID { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
    }

    public class EnvironmentalConcern
    {
        public int EnvironmentalConcernID { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
    }
    //3
    public class UpdatedStatus
    {
        public string UpdatedStatusID { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
    }

    public class CaseReport
    {
        public int CaseReportID { get; set; }
        public int UserInformationID { get; set; }
        public string FamilyName { get; set; }
        public string GivenName { get; set; }
        public string MaidenName { get; set; }
        public DateTime DateReported { get; set; }
        public int EnvironmentalConcernID { get; set; }
        public string Concern { get; set; }
        public int UpdatedStatusID { get; set; }
        public string UpdatedStatus { get; set; }
        public DateTime UpdatedStatusDate { get; set; }
        public string CaseLocation { get; set; }
        public string CaseReportPhoto { get; set; }
        public string XCoordinates { get; set; }
        public string YCoordinates { get; set; }

        public int VolunteerID { get; set; }

        public string Handler_GivenName { get; set; }
        public string Handler_MaidenName { get; set; }
        public string Handler_FamilyName { get; set; }
        public string Notes { get; set; }
    }

    public class UserInformation
    {
        public int UserInformationID { get; set; }
        public int UserTypeID { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string GivenName { get; set; }
        public string MaidenName { get; set; }
        public string FamilyName { get; set; }
        public string Email { get; set; }
        public string Notes { get; set; }
    }

    public class Volunteer
    {
        public int VolunteerID { get; set; }
        public string GivenName { get; set; }
        public string MaidenName { get; set; }
        public string FamilyName { get; set; }
        public string FullName { get; set; }
        public string Notes { get; set; }
    }

    //CHARTS
    public class CHART_MAKER
    {
        public int Number { get; set; }
        public string Description { get; set; }
    }



    //DASHBOARDS
    public class DASHBOARD_Home
    {
        public int Year { get; set; }
        public int EnvironmentalConcernID { get; set; }
    }


    public class AreaDetails
    {
        public int Number { get; set; }
        public string CaseLocation { get; set; }
        public int L_Submitted { get; set; }
        public int L_Rejected { get; set; }
        public int L_Accepted { get; set; }
        public int L_InProgress { get; set; }
        public int L_Completed { get; set; }

        public int W_Submitted { get; set; }
        public int W_Rejected { get; set; }
        public int W_Accepted { get; set; }
        public int W_InProgress { get; set; }
        public int W_Completed { get; set; }
    }

    public class Leaderboard
    {
        public int UserInformationID { get; set; }
        public int Points { get; set; }
        public string Area_1_Report { get; set; }
        public string Area_2_Report { get; set; }
        public string Area_3_Report { get; set; }
        public string Area_4_Report { get; set; }
        public string Area_5_Report { get; set; }

        public int Areas { get; set; }
        public string FullName { get; set; }
        public string Username { get; set; }

        public int LandPoints { get; set; }
        public string Land_1_Report { get; set; }
        public string Land_5_Reports { get; set; }
        public string Land_10_Reports { get; set; }
        public string Land_15_Reports { get; set; }
        public string Land_20_Reports { get; set; }
        public string Land_25_Reports { get; set; }
        public string Land_30_Reports { get; set; }
        public string Land_50_Reports { get; set; }

        public int WaterPoints { get; set; }
        public string Water_1_Report { get; set; }
        public string Water_5_Reports { get; set; }
        public string Water_10_Reports { get; set; }
        public string Water_15_Reports { get; set; }
        public string Water_20_Reports { get; set; }
        public string Water_25_Reports { get; set; }
        public string Water_30_Reports { get; set; }
        public string Water_50_Reports { get; set; }
        public string Overall_100_Reports { get; set; }
    }

}