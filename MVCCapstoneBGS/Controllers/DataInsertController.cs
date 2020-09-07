using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    public class DataInsertController : Controller
    {

        IDataProvider _IDataProvider;
        public DataInsertController()
        {
            _IDataProvider = new DataProvider();
        }

        public ActionResult Index()
        {
            return View();
        }

        public void InsertVolunteer(string GivenName, string MaidenName, string FamilyName)
        {
            _IDataProvider.InsertVolunteer(GivenName,MaidenName,FamilyName);
        }


        public void CreateCommunityUser(string UserName, string Password, string Email, string GivenName, string MaidenName, string FamilyName)
        {

            //string page = Request.Url.AbsolutePath.ToString();
            ////https://localhost:44314/DataInsert/CreateCommunityUser?Username=Edward31313&Password=131313&Email=Hello@yahoo.com&GivenName=Edwardo&MaidenName=Sanglay&FamilyName=Garcia
            //var data = _IDataProvider.InsertUserInformation(2, UserName, Password, Email, GivenName, MaidenName, FamilyName);
            //return Json(data,JsonRequestBehavior.AllowGet);
            _IDataProvider.InsertUserInformation(2, UserName, Password, Email, GivenName, MaidenName, FamilyName);


        }

        public void CreateAdministrator(string UserName, string Password, string Email, string GivenName, string MaidenName, string FamilyName)
        {
            //https://localhost:44314/DataInsert/CreateCommunityUser?Username=Edward31313&Password=131313&Email=Hello@yahoo.com&GivenName=Edwardo&MaidenName=Sanglay&FamilyName=Garcia
            _IDataProvider.InsertUserInformation(1, UserName, Password, Email, GivenName, MaidenName, FamilyName);
        }
    }
}