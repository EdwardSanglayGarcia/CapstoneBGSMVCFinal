using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCCapstoneBGS
{
    using System.Data.SqlClient;
    using System.Configuration;

    public class DataAccess
    {
        protected static string constring = ConfigurationManager.ConnectionStrings["CapstoneDemo"].ConnectionString;

        protected static SqlConnection con;
        protected static SqlCommand cmd;
        protected static SqlDataAdapter da;
        protected static SqlDataReader dr;
        protected static SqlTransaction trx;
    }
}