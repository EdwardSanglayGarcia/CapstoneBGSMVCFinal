using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCCapstoneBGS
{
    public struct LabelStruct
    {

        public struct Administrator
        {
            public const string AdministratorHomepage = "Administrator";
            public const string Submitted = "Submitted";
            public const string Accepted = "Accepted";
            public const string InProgress = "In Progress";
            public const string Completed = "Completed";
            public const string Rejected = "Rejected";
            public const string MonthlyReports = "Monthly Reports";
            public const string YearlyReports = "Yearly Reports";
            public const string Twitter = "Twitter";
            public const string Volunteers = "Volunteers";
        }

        public struct CommunityUser
        {
            public const string CommunityUserHomepage = "Community User";
            public const string SubmitReport = "Submit Report";
            public const string ViewStatus = "View Status";
            public const string Achievements = "Achievements";
        }

        public struct Others
        {
            public const string Leaderboards = "Leaderboards";

        }


    }
}