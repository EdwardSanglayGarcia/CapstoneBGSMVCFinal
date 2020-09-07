using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCCapstoneBGS.Controllers
{
    using TweetSharp;

    using OfficeOpenXml;
    using OfficeOpenXml.Drawing;
    using OfficeOpenXml.Style;
    using System.Drawing;

    using System.IO;
    using System.Web;
    using System.Net;
    using System.Globalization;

    public class DataGenerateController : Controller
    {

        IDataProvider _IDataProvider;
        public DataGenerateController()
        {
            _IDataProvider = new DataProvider();
        }


        // GET: DataGenerate
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult TestTwitter(string message)
        {
            DefaultData cmd = new DefaultData();

            var service = new TwitterService(cmd.CONSUMER_KEY, cmd.CONSUMER_SECRET);
            service.AuthenticateWith(cmd.ACCESS_TOKEN, cmd.ACCESS_TOKEN_SECRET);
            TwitterStatus result = service.SendTweet(new SendTweetOptions
            {
                Status = message + " #TerraTechPH"
            });

            return View("~/Views/Entities/Twitter.cshtml");
        }

        public void TweetSomething(string message)
        {
            string image;
            image = @"C:\Users\pc\Desktop\Capstone Codes\Web Application\MVCCapstoneBGS\TerraMaster\upload\shutterstock_124871620.jpg";
            //int imageID=0;
            //List<string> imageList = new List<string>();
            //imageList.Add(@"C:\Users\pc\Desktop\Anime\Victorique.png");
            //imageList.Add(@"C:\Users\pc\Desktop\Anime\Shidou.png");

            //string IMAGE = @"C:\Users\pc\Desktop\Anime\Shidou.png";

            DefaultData cmd = new DefaultData();

            var service = new TwitterService(cmd.CONSUMER_KEY, cmd.CONSUMER_SECRET);
            service.AuthenticateWith(cmd.ACCESS_TOKEN, cmd.ACCESS_TOKEN_SECRET);

            using (var stream = new FileStream(image, FileMode.Open))
            {
                service.SendTweetWithMedia(new SendTweetWithMediaOptions
                {
                    Status = message + " #TerraTechPH",
                    Images = new Dictionary<string, Stream> { { image, stream } }
                });
            }
        }

        public ActionResult GenerateYearReport(int year)
        {
            var memoryStream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DateTimeFormatInfo mfi = new DateTimeFormatInfo();
            using (var excelPackage = new ExcelPackage(memoryStream))
            {

                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "VDWWD";
                excelPackage.Workbook.Properties.Title = "Title of Document";
                excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                #region 
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("Summarized " + year + " Report");
                try
                {

                    int height = 230;
                    int width = 230;

                    string DENRLogo = @"C:\Users\pc\Desktop\DENRLogo.png";
                    Image DENRLOGO = Image.FromFile(DENRLogo);
                    ExcelPicture picDENR = ws.Drawings.AddPicture("DENR", DENRLOGO);
                    picDENR.SetPosition(0, 0, 0, 0);//5,3
                    picDENR.SetSize(height, width);

                    string PHLogo = @"C:\Users\pc\Desktop\PHLogo.png";
                    Image PHILIPPINELOGO = Image.FromFile(PHLogo);
                    ExcelPicture picPH = ws.Drawings.AddPicture("PH", PHILIPPINELOGO);
                    picPH.SetPosition(0, 0, 10, 0);//5,3
                    picPH.SetSize(height, width);


                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw;
                }
                //Resize the columns A-M
                ws.Column(1).Width = 24;
                for (int startColumns = 2; startColumns <= 13; startColumns++)
                {
                    ws.Column(startColumns).Width = 12;
                }

                //Resize the row 1-7
                ws.Row(1).Height = 40;
                ws.Row(2).Height = 35;
                ws.Row(3).Height = 25;
                ws.Row(4).Height = 25;
                ws.Row(5).Height = 20;
                ws.Row(6).Height = 20;

                //Merging
                for (int mergeHeader = 1; mergeHeader <= 7; mergeHeader++)
                {
                    ws.Cells[mergeHeader, 1, mergeHeader, 13].Merge = true;
                }

                //Adding Descriptive Values
                ws.Cells[1, 1, 1, 13].Value = "Republic of the Philippines";
                ws.Cells[2, 1, 2, 13].Value = "Department of Environment and National Resources";
                ws.Cells[3, 1, 3, 13].Value = "Environmental Management Bureau";
                ws.Cells[4, 1, 4, 13].Value = "Environmental Monitoring and Enforcement Division";
                ws.Cells[5, 1, 5, 13].Value = "National Ecology Center Compound East Avenue Diliman Quezon City";
                ws.Cells[6, 1, 6, 13].Value = "Tel/Email: 8931-3506 | 8931-2684 | RecordsNCR@emb.gov.ph";

                //Changing Font Size
                ws.Cells[1, 1, 1, 13].Style.Font.Size = 30;
                ws.Cells[2, 1, 2, 13].Style.Font.Size = 25;
                ws.Cells[3, 1, 3, 13].Style.Font.Size = 20;
                ws.Cells[4, 1, 4, 13].Style.Font.Size = 20;
                ws.Cells[5, 1, 5, 13].Style.Font.Size = 16;
                ws.Cells[6, 1, 6, 13].Style.Font.Size = 16;

                //Change the Font Style of the Header
                ws.Cells[1, 1, 6, 13].Style.Font.Name = "Times New Roman";

                //Align to center the header
                ws.Cells[1, 1, 6, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[1, 1, 6, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //Secondary Header for Data
                ws.Cells[8, 1, 9, 13].Merge = true;
                ws.Cells[8, 1, 9, 13].Value = "Summarized Yearly " + year + " Report";
                ws.Cells[8, 1, 9, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[8, 1, 9, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[8, 1, 9, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[8, 1, 9, 13].Style.Font.Name = "Arial Narrow";
                ws.Cells[8, 1, 9, 13].Style.Font.Bold = true;
                ws.Cells[8, 1, 9, 13].Style.Font.Size = 15;

                //Main Header
                ws.Cells[10, 1, 11, 1].Merge = true;
                ws.Cells[10, 1, 11, 1].Value = "Month";
                ws.Cells[10, 1, 11, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[10, 1, 11, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[10, 1, 11, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[10, 1, 11, 1].Style.Font.Name = "Arial Narrow";
                ws.Cells[10, 1, 11, 1].Style.Font.Bold = true;
                ws.Cells[10, 1, 11, 1].Style.Font.Size = 15;

                for (int dataHeader = 2; dataHeader <= 13; dataHeader++)
                {
                    ws.Cells[11, dataHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[11, dataHeader].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[11, dataHeader].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[11, dataHeader].Style.Font.Name = "Arial Narrow";
                    ws.Cells[11, dataHeader].Style.Font.Size = 12;
                }

                //L - W - TOTAL + header

                for (int L = 2; L <= 13; L += 3)
                {
                    ws.Cells[11, L].Value = "L";
                    ws.Cells[11, L + 1].Value = "W";
                    ws.Cells[11, L + 2].Value = "TOTAL";
                    ws.Cells[10, L, 10, L + 2].Merge = true;
                    ws.Cells[10, L, 10, L + 2].Style.Border.BorderAround(ExcelBorderStyle.Double);
                }
                ws.Cells[10, 2, 11, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[10, 2, 11, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[10, 2, 11, 13].Style.Font.Bold = true;
                ws.Cells[10, 2, 11, 13].Style.Font.Name = "Arial Narrow";
                ws.Cells[10, 2, 11, 13].Style.Font.Size = 12;
                ws.Cells[10, 2].Value = "Submitted";
                ws.Cells[10, 5].Value = "Rejected";
                ws.Cells[10, 8].Value = "Accepted";
                ws.Cells[10, 11].Value = "Completed";

                //DATA SECTION

                ws.Cells[12, 1].Value = "January";
                ws.Cells[13, 1].Value = "February";
                ws.Cells[14, 1].Value = "March";
                ws.Cells[15, 1].Value = "April";
                ws.Cells[16, 1].Value = "May";
                ws.Cells[17, 1].Value = "June";
                ws.Cells[18, 1].Value = "July";
                ws.Cells[19, 1].Value = "August";
                ws.Cells[20, 1].Value = "September";
                ws.Cells[21, 1].Value = "October";
                ws.Cells[22, 1].Value = "November";
                ws.Cells[23, 1].Value = "December";
                ws.Cells[24, 1].Value = "TOTAL";

                ws.Cells[12, 1, 24, 1].Style.Font.Bold = true;

                for (int rowData = 12; rowData <= 24; rowData++)
                {
                    for (int colData = 1; colData <= 13; colData++)
                    {
                        ws.Cells[rowData, colData].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[rowData, colData].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[rowData, colData].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[rowData, colData].Style.Font.Name = "Arial Narrow";
                        ws.Cells[rowData, colData].Style.Font.Size = 12;
                    }
                }

                for (int rowDataColor = 11; rowDataColor <= 24; rowDataColor++)
                {
                    for (int LWH = 2; LWH <= 13; LWH += 3)
                    {
                        //L
                        ws.Cells[rowDataColor, LWH].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[rowDataColor, LWH].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        //W
                        ws.Cells[rowDataColor, LWH + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[rowDataColor, LWH + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        //H
                        ws.Cells[rowDataColor, LWH + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[rowDataColor, LWH + 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218));
                    }
                }

                int DataRowStart = 12;
                //DATA SECTION
                for (int mainData = 1; mainData <= 12; mainData++)
                {
                    foreach (var data in _IDataProvider.GetMonthlyTotals(mainData, year))
                    {
                        Console.WriteLine(mainData + " - " + year);
                        ws.Cells[DataRowStart, 2].Value = data.L_Submitted;
                        ws.Cells[DataRowStart, 3].Value = data.W_Submitted;
                        ws.Cells[DataRowStart, 4].Formula = "SUM(" + ws.Cells[DataRowStart, 2] + ":" + ws.Cells[DataRowStart, 3] + ")";

                        ws.Cells[DataRowStart, 5].Value = data.L_Rejected;
                        ws.Cells[DataRowStart, 6].Value = data.W_Rejected;
                        ws.Cells[DataRowStart, 7].Formula = "SUM(" + ws.Cells[DataRowStart, 5] + ":" + ws.Cells[DataRowStart, 6] + ")";

                        ws.Cells[DataRowStart, 8].Value = data.L_Accepted;
                        ws.Cells[DataRowStart, 9].Value = data.W_Accepted;
                        ws.Cells[DataRowStart, 10].Formula = "SUM(" + ws.Cells[DataRowStart, 8] + ":" + ws.Cells[DataRowStart, 9] + ")";

                        ws.Cells[DataRowStart, 11].Value = data.L_Completed;
                        ws.Cells[DataRowStart, 12].Value = data.W_Completed;
                        ws.Cells[DataRowStart, 13].Formula = "SUM(" + ws.Cells[DataRowStart, 11] + ":" + ws.Cells[DataRowStart, 12] + ")";
                        DataRowStart++;
                    }
                }
                int totalRowStart = ws.Dimension.End.Row;

                for (int totalMainData = 2; totalMainData <= 13; totalMainData++)
                {
                    ws.Cells[totalRowStart, totalMainData].Formula = "SUM(" + ws.Cells[totalRowStart - 1, totalMainData] + ":" + ws.Cells[totalRowStart - 12, totalMainData] + ")";
                }

                int marginMake = ws.Dimension.End.Row + 2;

                ws.Cells[marginMake, 1, marginMake + 1, 13].Merge = true;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Font.Name = "Arial Narrow";
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Font.Size = 15;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Font.Bold = true;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Value = year + " Summarized Report per City / Municipality";

                int subDataHeaderStart = ws.Dimension.End.Row + 2;

                //  Console.WriteLine(ws.Dimension.End.Row); //27
                //START OF SUB DATA


                for (int repeatMonth = 1; repeatMonth <= 12; repeatMonth++)
                {
                    // ws.Cells[subDataHeaderStart + 1, 1].Value = "Record for month num" +repeatMonth;

                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Merge = true;
                    ws.Cells[subDataHeaderStart, 2, subDataHeaderStart, 4].Merge = true;
                    ws.Cells[subDataHeaderStart, 5, subDataHeaderStart, 7].Merge = true;
                    ws.Cells[subDataHeaderStart, 8, subDataHeaderStart, 10].Merge = true;
                    ws.Cells[subDataHeaderStart, 11, subDataHeaderStart, 13].Merge = true;

                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart, 2, subDataHeaderStart, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart, 5, subDataHeaderStart, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart, 8, subDataHeaderStart, 10].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart, 11, subDataHeaderStart, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Value = mfi.GetMonthName(repeatMonth).ToString() + " Breakdown City/Municipality";
                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Style.WrapText = true;

                    ws.Cells[subDataHeaderStart, 2, subDataHeaderStart, 4].Value = "Submitted";
                    ws.Cells[subDataHeaderStart, 5, subDataHeaderStart, 7].Value = "Rejected";
                    ws.Cells[subDataHeaderStart, 8, subDataHeaderStart, 10].Value = "Accepted";
                    ws.Cells[subDataHeaderStart, 11, subDataHeaderStart, 13].Value = "Completed";

                    for (int L = 2; L <= 13; L += 3)
                    {
                        ws.Cells[subDataHeaderStart + 1, L].Value = "L";
                        ws.Cells[subDataHeaderStart + 1, L].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[subDataHeaderStart + 1, L].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        ws.Cells[subDataHeaderStart + 1, L + 1].Value = "W";
                        ws.Cells[subDataHeaderStart + 1, L + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[subDataHeaderStart + 1, L + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        ws.Cells[subDataHeaderStart + 1, L + 2].Value = "TOTAL";
                        ws.Cells[subDataHeaderStart + 1, L + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[subDataHeaderStart + 1, L + 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218));

                        ws.Cells[subDataHeaderStart + 1, L].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 1, L + 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 1, L + 2].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    }

                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Font.Name = "Arial Narrow";
                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Font.Size = 12;
                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Font.Bold = true;

                    foreach (var getData in _IDataProvider.GetAreaDetailsPerMonthYear(repeatMonth, year))
                    {
                        for (int L = 2; L <= 13; L += 3)
                        {
                            ws.Cells[subDataHeaderStart + 2, L].Value = "L";
                            ws.Cells[subDataHeaderStart + 2, L].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[subDataHeaderStart + 2, L].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));

                            ws.Cells[subDataHeaderStart + 2, L + 1].Value = "W";
                            ws.Cells[subDataHeaderStart + 2, L + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[subDataHeaderStart + 2, L + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));

                            ws.Cells[subDataHeaderStart + 2, L + 2].Value = "TOTAL";
                            ws.Cells[subDataHeaderStart + 2, L + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[subDataHeaderStart + 2, L + 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218));
                        }

                        ws.Cells[subDataHeaderStart + 2, 1].Value = getData.CaseLocation;
                        ws.Cells[subDataHeaderStart + 2, 2].Value = getData.L_Submitted;
                        ws.Cells[subDataHeaderStart + 2, 3].Value = getData.W_Submitted;
                        ws.Cells[subDataHeaderStart + 2, 4].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 2] + ":" + ws.Cells[subDataHeaderStart + 2, 3] + ")";

                        ws.Cells[subDataHeaderStart + 2, 5].Value = getData.L_Rejected;
                        ws.Cells[subDataHeaderStart + 2, 6].Value = getData.W_Rejected;
                        ws.Cells[subDataHeaderStart + 2, 7].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 5] + ":" + ws.Cells[subDataHeaderStart + 2, 6] + ")";

                        ws.Cells[subDataHeaderStart + 2, 8].Value = getData.L_Accepted;
                        ws.Cells[subDataHeaderStart + 2, 9].Value = getData.W_Accepted;
                        ws.Cells[subDataHeaderStart + 2, 10].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 8] + ":" + ws.Cells[subDataHeaderStart + 2, 9] + ")";

                        ws.Cells[subDataHeaderStart + 2, 11].Value = getData.L_Completed;
                        ws.Cells[subDataHeaderStart + 2, 12].Value = getData.W_Completed;
                        ws.Cells[subDataHeaderStart + 2, 13].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 11] + ":" + ws.Cells[subDataHeaderStart + 2, 12] + ")";
                        ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.Font.Name = "Arial Narrow";
                        ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.Font.Size = 12;
                        ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                        ws.Cells[subDataHeaderStart + 2, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        ws.Cells[subDataHeaderStart + 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        ws.Cells[subDataHeaderStart + 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 6].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        ws.Cells[subDataHeaderStart + 2, 8].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 10].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        ws.Cells[subDataHeaderStart + 2, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 12].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        ws.Cells[subDataHeaderStart + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        subDataHeaderStart++;
                    }


                    subDataHeaderStart += 5;


                }
                #endregion

                #region LandSpreadSheet

                for (int loopingMonth = 1; loopingMonth <= 12; loopingMonth++)
                {
                    ExcelWorksheet wsL = excelPackage.Workbook.Worksheets.Add(mfi.GetMonthName(loopingMonth)+" " + year + " Land Report");

                    try
                    {
                        int height = 230;
                        int width = 230;

                        string DENRLogo = @"C:\Users\pc\Desktop\DENRLogo.png";
                        Image DENRLOGO = Image.FromFile(DENRLogo);
                        ExcelPicture picDENR = wsL.Drawings.AddPicture("DENR", DENRLOGO);
                        picDENR.SetPosition(0, 0, 0, 0);//5,3
                        picDENR.SetSize(height, width);

                        string PHLogo = @"C:\Users\pc\Desktop\PHLogo.png";
                        Image PHILIPPINELOGO = Image.FromFile(PHLogo);
                        ExcelPicture picPH = wsL.Drawings.AddPicture("PH", PHILIPPINELOGO);
                        picPH.SetPosition(0, 0, 10, 0);//5,3
                        picPH.SetSize(height, width);
                    }

                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        throw;
                    }


                    //Resize the columns A-M
                    wsL.Column(1).Width = 24;
                    for (int startColumns = 2; startColumns <= 13; startColumns++)
                    {
                        wsL.Column(startColumns).Width = 12;
                    }

                    //Resize the row 1-7
                    wsL.Row(1).Height = 40;
                    wsL.Row(2).Height = 35;
                    wsL.Row(3).Height = 25;
                    wsL.Row(4).Height = 25;
                    wsL.Row(5).Height = 20;
                    wsL.Row(6).Height = 20;

                    //Merging
                    for (int mergeHeader = 1; mergeHeader <= 7; mergeHeader++)
                    {
                        wsL.Cells[mergeHeader, 1, mergeHeader, 13].Merge = true;
                    }

                    //Adding Descriptive Values
                    wsL.Cells[1, 1, 1, 13].Value = "Republic of the Philippines";
                    wsL.Cells[2, 1, 2, 13].Value = "Department of Environment and National Resources";
                    wsL.Cells[3, 1, 3, 13].Value = "Environmental Management Bureau";
                    wsL.Cells[4, 1, 4, 13].Value = "Environmental Monitoring and Enforcement Division";
                    wsL.Cells[5, 1, 5, 13].Value = "National Ecology Center Compound East Avenue Diliman Quezon City";
                    wsL.Cells[6, 1, 6, 13].Value = "Tel/Email: 8931-3506 | 8931-2684 | RecordsNCR@emb.gov.ph";

                    //Changing Font Size
                    wsL.Cells[1, 1, 1, 13].Style.Font.Size = 30;
                    wsL.Cells[2, 1, 2, 13].Style.Font.Size = 25;
                    wsL.Cells[3, 1, 3, 13].Style.Font.Size = 20;
                    wsL.Cells[4, 1, 4, 13].Style.Font.Size = 20;
                    wsL.Cells[5, 1, 5, 13].Style.Font.Size = 16;
                    wsL.Cells[6, 1, 6, 13].Style.Font.Size = 16;

                    //Change the Font Style of the Header
                    wsL.Cells[1, 1, 6, 13].Style.Font.Name = "Times New Roman";

                    //Align to center the header
                    wsL.Cells[1, 1, 6, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[1, 1, 6, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    wsL.Cells[8, 1, 9, 13].Merge = true;
                    wsL.Cells[8, 1, 9, 13].Value = mfi.GetMonthName(loopingMonth) + " " + year + " Report";
                    wsL.Cells[8, 1, 9, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[8, 1, 9, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[8, 1, 9, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[8, 1, 9, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[8, 1, 9, 13].Style.Font.Size = 15;

                    int SubmittedHeader = wsL.Dimension.End.Row + 1;
                    #region Submit
                    wsL.Cells[SubmittedHeader+1, 1, SubmittedHeader + 2, 13].Merge = true;
                    wsL.Cells[SubmittedHeader+1, 1, SubmittedHeader + 2, 13].Value = "Submitted Reports";
                    wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.Font.Size = 15;

                   
                    wsL.Cells[SubmittedHeader +3, 2, SubmittedHeader + 3, 3].Merge = true;
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Merge = true;
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Merge = true;
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Merge = true;
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Merge = true;
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Merge = true;

                    wsL.Cells[SubmittedHeader + 3, 1].Value = "Reference No.";

                    wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Value = "Username";
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Value = "Date Reported";
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Value = "Location";
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Value = "Coordinates";
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Value = "Date Updated";
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Value = "Handler";

                    //REFERENCE NO
                    wsL.Cells[SubmittedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 1].Style.Font.Size = 12;

                    //USERNAME
                    wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.Font.Size = 12;

                    //DATE REPORTED
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.Font.Size = 12;

                    //LOCATION
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.Font.Size = 12;

                    //COORDINATES
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.Font.Size = 12;

                    //DATE UPDATED
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.Font.Size = 12;

                    //HANDLER
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.Font.Size = 12;

                    int SData = wsL.Dimension.End.Row + 1;
                    #endregion
                    //SUBMITTED DATA SECTION
                    foreach (var SLandSubmit in _IDataProvider.GetCaseReport(1).Where(x=>x.UpdatedStatusDate.Year==year && x.UpdatedStatusDate.Month==loopingMonth).OrderBy(o=>o.EnvironmentalConcernID).ToList())
                    {
                        if (SLandSubmit.EnvironmentalConcernID == 1)
                        {
                            wsL.Cells[SData, 1, SData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[SData, 1, SData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        }
                        else
                        {
                            wsL.Cells[SData, 1, SData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[SData, 1, SData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        }

                        #region DesignSubmit
                        wsL.Cells[SData, 1].Merge = true;
                        wsL.Cells[SData, 2, SData, 3].Merge = true;
                        wsL.Cells[SData, 4, SData, 5].Merge = true;
                        wsL.Cells[SData, 6, SData, 7].Merge = true;
                        wsL.Cells[SData, 8, SData, 9].Merge = true;
                        wsL.Cells[SData, 10, SData, 11].Merge = true;
                        wsL.Cells[SData, 12, SData, 13].Merge = true;

                        wsL.Cells[SData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[SData, 2, SData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[SData, 4, SData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[SData, 6, SData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[SData, 8, SData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[SData, 10, SData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[SData, 12, SData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        wsL.Cells[SData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[SData, 2, SData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[SData, 4, SData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[SData, 6, SData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[SData, 8, SData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[SData, 10, SData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[SData, 12, SData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        wsL.Cells[SData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[SData, 2, SData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[SData, 4, SData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[SData, 6, SData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[SData, 8, SData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[SData, 10, SData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[SData, 12, SData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        wsL.Cells[SData, 1].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[SData, 2, SData, 3].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[SData, 4, SData, 5].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[SData, 6, SData, 7].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[SData, 8, SData, 9].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[SData, 10, SData, 11].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[SData, 12, SData, 13].Style.Font.Name = "Arial Narrow";


                        wsL.Cells[SData, 1].Style.Font.Size = 12;
                        wsL.Cells[SData, 2, SData, 3].Style.Font.Size = 12;
                        wsL.Cells[SData, 4, SData, 5].Style.Font.Size = 12;
                        wsL.Cells[SData, 6, SData, 7].Style.Font.Size = 12;
                        wsL.Cells[SData, 8, SData, 9].Style.Font.Size = 12;
                        wsL.Cells[SData, 10, SData, 11].Style.Font.Size = 12;
                        wsL.Cells[SData, 12, SData, 13].Style.Font.Size = 12;
                        #endregion

                        wsL.Cells[SData, 1].Value = SLandSubmit.CaseReportID;
                        wsL.Cells[SData, 2, SData, 3].Value = SLandSubmit.FamilyName + ", " + SLandSubmit.GivenName + " " + SLandSubmit.MaidenName;
                        wsL.Cells[SData, 4, SData, 5].Value = SLandSubmit.DateReported.ToString();
                        wsL.Cells[SData, 6, SData, 7].Value = SLandSubmit.CaseLocation;
                        wsL.Cells[SData, 8, SData, 9].Value = SLandSubmit.XCoordinates + "," + SLandSubmit.YCoordinates;
                        wsL.Cells[SData, 10, SData, 11].Value = SLandSubmit.UpdatedStatusDate.ToString();
                        wsL.Cells[SData, 12, SData, 13].Value = SLandSubmit.Handler_FamilyName + ", " + SLandSubmit.Handler_GivenName + " " + SLandSubmit.Handler_MaidenName;

                        SData++;
                    }

                    int AcceptedHeader = wsL.Dimension.End.Row + 1;
                    #region Accept
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Merge = true;
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Value = "Accepted Reports";
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.Font.Size = 15;


                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Merge = true;
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Merge = true;
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Merge = true;
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Merge = true;
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Merge = true;
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Merge = true;

                    wsL.Cells[AcceptedHeader + 3, 1].Value = "Reference No.";

                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Value = "Username";
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Value = "Date Reported";
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Value = "Location";
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Value = "Coordinates";
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Value = "Date Updated";
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Value = "Handler";

                    //REFERENCE NO
                    wsL.Cells[AcceptedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 1].Style.Font.Size = 12;

                    //USERNAME
                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.Font.Size = 12;

                    //DATE REPORTED
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.Font.Size = 12;

                    //LOCATION
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.Font.Size = 12;

                    //COORDINATES
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.Font.Size = 12;

                    //DATE UPDATED
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.Font.Size = 12;

                    //HANDLER
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.Font.Size = 12;

                    int AData = wsL.Dimension.End.Row + 1;

                    #endregion
                    //ACCEPTED DATA SECTION
                    foreach (var SLandAccept in _IDataProvider.GetCaseReport(3).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == loopingMonth).ToList())
                    {
                        if (SLandAccept.EnvironmentalConcernID == 1)
                        {
                            wsL.Cells[AData, 1, AData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[AData, 1, AData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        }
                        else
                        {
                            wsL.Cells[AData, 1, AData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[AData, 1, AData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        }

                        #region DesignSubmit
                        wsL.Cells[AData, 1].Merge = true;
                        wsL.Cells[AData, 2, AData, 3].Merge = true;
                        wsL.Cells[AData, 4, AData, 5].Merge = true;
                        wsL.Cells[AData, 6, AData, 7].Merge = true;
                        wsL.Cells[AData, 8, AData, 9].Merge = true;
                        wsL.Cells[AData, 10, AData, 11].Merge = true;
                        wsL.Cells[AData, 12, AData, 13].Merge = true;

                        wsL.Cells[AData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[AData, 2, AData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[AData, 4, AData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[AData, 6, AData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[AData, 8, AData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[AData, 10, AData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[AData, 12, AData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        wsL.Cells[AData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[AData, 2, AData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[AData, 4, AData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[AData, 6, AData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[AData, 8, AData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[AData, 10, AData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[AData, 12, AData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        wsL.Cells[AData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[AData, 2, AData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[AData, 4, AData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[AData, 6, AData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[AData, 8, AData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[AData, 10, AData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[AData, 12, AData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        wsL.Cells[AData, 1].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[AData, 2, AData, 3].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[AData, 4, AData, 5].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[AData, 6, AData, 7].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[AData, 8, AData, 9].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[AData, 10, AData, 11].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[AData, 12, AData, 13].Style.Font.Name = "Arial Narrow";


                        wsL.Cells[AData, 1].Style.Font.Size = 12;
                        wsL.Cells[AData, 2, AData, 3].Style.Font.Size = 12;
                        wsL.Cells[AData, 4, AData, 5].Style.Font.Size = 12;
                        wsL.Cells[AData, 6, AData, 7].Style.Font.Size = 12;
                        wsL.Cells[AData, 8, AData, 9].Style.Font.Size = 12;
                        wsL.Cells[AData, 10, AData, 11].Style.Font.Size = 12;
                        wsL.Cells[AData, 12, AData, 13].Style.Font.Size = 12;
                        #endregion

                        wsL.Cells[AData, 1].Value = SLandAccept.CaseReportID;
                        wsL.Cells[AData, 2, AData, 3].Value = SLandAccept.FamilyName + ", " + SLandAccept.GivenName + " " + SLandAccept.MaidenName;
                        wsL.Cells[AData, 4, AData, 5].Value = SLandAccept.DateReported.ToString();
                        wsL.Cells[AData, 6, AData, 7].Value = SLandAccept.CaseLocation;
                        wsL.Cells[AData, 8, AData, 9].Value = SLandAccept.XCoordinates + "," + SLandAccept.YCoordinates;
                        wsL.Cells[AData, 10, AData, 11].Value = SLandAccept.UpdatedStatusDate.ToString();
                        wsL.Cells[AData, 12, AData, 13].Value = SLandAccept.Handler_FamilyName + ", " + SLandAccept.Handler_GivenName + " " + SLandAccept.Handler_MaidenName;

                        AData++;
                    }

                    int RejectedHeader = wsL.Dimension.End.Row + 1;
                    #region Reject
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Merge = true;
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Value = "Rejected Reports";
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.Font.Size = 15;


                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Merge = true;
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Merge = true;
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Merge = true;
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Merge = true;
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Merge = true;
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Merge = true;

                    wsL.Cells[RejectedHeader + 3, 1].Value = "Reference No.";

                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Value = "Username";
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Value = "Date Reported";
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Value = "Location";
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Value = "Coordinates";
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Value = "Date Updated";
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Value = "Handler";

                    //REFERENCE NO
                    wsL.Cells[RejectedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 1].Style.Font.Size = 12;

                    //USERNAME
                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.Font.Size = 12;

                    //DATE REPORTED
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.Font.Size = 12;

                    //LOCATION
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.Font.Size = 12;

                    //COORDINATES
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.Font.Size = 12;

                    //DATE UPDATED
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.Font.Size = 12;

                    //HANDLER
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.Font.Size = 12;

                    int RData = wsL.Dimension.End.Row + 1;
                    #endregion
                    //REJECTED DATA SECTION
                    foreach (var LLandReject in _IDataProvider.GetCaseReport(2).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == loopingMonth).ToList())
                    {
                        #region DesignSubmit
                        wsL.Cells[RData, 1].Merge = true;
                        wsL.Cells[RData, 2, RData, 3].Merge = true;
                        wsL.Cells[RData, 4, RData, 5].Merge = true;
                        wsL.Cells[RData, 6, RData, 7].Merge = true;
                        wsL.Cells[RData, 8, RData, 9].Merge = true;
                        wsL.Cells[RData, 10, RData, 11].Merge = true;
                        wsL.Cells[RData, 12, RData, 13].Merge = true;

                        wsL.Cells[RData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[RData, 2, RData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[RData, 4, RData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[RData, 6, RData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[RData, 8, RData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[RData, 10, RData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[RData, 12, RData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        wsL.Cells[RData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[RData, 2, RData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[RData, 4, RData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[RData, 6, RData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[RData, 8, RData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[RData, 10, RData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[RData, 12, RData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        wsL.Cells[RData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[RData, 2, RData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[RData, 4, RData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[RData, 6, RData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[RData, 8, RData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[RData, 10, RData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[RData, 12, RData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        wsL.Cells[RData, 1].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[RData, 2, RData, 3].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[RData, 4, RData, 5].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[RData, 6, RData, 7].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[RData, 8, RData, 9].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[RData, 10, RData, 11].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[RData, 12, RData, 13].Style.Font.Name = "Arial Narrow";


                        wsL.Cells[RData, 1].Style.Font.Size = 12;
                        wsL.Cells[RData, 2, RData, 3].Style.Font.Size = 12;
                        wsL.Cells[RData, 4, RData, 5].Style.Font.Size = 12;
                        wsL.Cells[RData, 6, RData, 7].Style.Font.Size = 12;
                        wsL.Cells[RData, 8, RData, 9].Style.Font.Size = 12;
                        wsL.Cells[RData, 10, RData, 11].Style.Font.Size = 12;
                        wsL.Cells[RData, 12, RData, 13].Style.Font.Size = 12;
                        #endregion

                        if (LLandReject.EnvironmentalConcernID == 1)
                        {
                            wsL.Cells[RData, 1, RData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[RData, 1, RData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        }
                        else
                        {
                            wsL.Cells[RData, 1, RData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[RData, 1, RData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        }

                        wsL.Cells[RData, 1].Value = LLandReject.CaseReportID;
                        wsL.Cells[RData, 2, RData, 3].Value = LLandReject.FamilyName + ", " + LLandReject.GivenName + " " + LLandReject.MaidenName;
                        wsL.Cells[RData, 4, RData, 5].Value = LLandReject.DateReported.ToString();
                        wsL.Cells[RData, 6, RData, 7].Value = LLandReject.CaseLocation;
                        wsL.Cells[RData, 8, RData, 9].Value = LLandReject.XCoordinates + "," + LLandReject.YCoordinates;
                        wsL.Cells[RData, 10, RData, 11].Value = LLandReject.UpdatedStatusDate.ToString();
                        wsL.Cells[RData, 12, RData, 13].Value = LLandReject.Handler_FamilyName + ", " + LLandReject.Handler_GivenName + " " + LLandReject.Handler_MaidenName;

                        RData++;
                    }
                   
                    int InProgressHeader = wsL.Dimension.End.Row + 1;
                    #region InProgress
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Merge = true;
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Value = "In Progress Reports";
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.Font.Size = 15;


                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Merge = true;
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Merge = true;
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Merge = true;
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Merge = true;
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Merge = true;
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Merge = true;

                    wsL.Cells[InProgressHeader + 3, 1].Value = "Reference No.";

                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Value = "Username";
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Value = "Date Reported";
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Value = "Location";
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Value = "Coordinates";
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Value = "Date Updated";
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Value = "Handler";

                    //REFERENCE NO
                    wsL.Cells[InProgressHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 1].Style.Font.Size = 12;

                    //USERNAME
                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.Font.Size = 12;

                    //DATE REPORTED
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.Font.Size = 12;

                    //LOCATION
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.Font.Size = 12;

                    //COORDINATES
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.Font.Size = 12;

                    //DATE UPDATED
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.Font.Size = 12;

                    //HANDLER
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.Font.Size = 12;

                    int IPData = wsL.Dimension.End.Row + 1;
                    #endregion
                    //  IN PROGRESS DATA SECTION
                    foreach (var IPLand in _IDataProvider.GetCaseReport(4).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == loopingMonth).ToList())
                    {
                        #region DesignSubmit

                        if (IPLand.EnvironmentalConcernID == 1)
                        {
                            wsL.Cells[IPData, 1, IPData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[IPData, 1, IPData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        }
                        else
                        {
                            wsL.Cells[IPData, 1, IPData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[IPData, 1, IPData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        }


                        wsL.Cells[IPData, 1].Merge = true;
                        wsL.Cells[IPData, 2, IPData, 3].Merge = true;
                        wsL.Cells[IPData, 4, IPData, 5].Merge = true;
                        wsL.Cells[IPData, 6, IPData, 7].Merge = true;
                        wsL.Cells[IPData, 8, IPData, 9].Merge = true;
                        wsL.Cells[IPData, 10, IPData, 11].Merge = true;
                        wsL.Cells[IPData, 12, IPData, 13].Merge = true;

                        wsL.Cells[IPData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[IPData, 2, IPData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[IPData, 4, IPData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[IPData, 6, IPData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[IPData, 8, IPData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[IPData, 10, IPData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[IPData, 12, IPData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        wsL.Cells[IPData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[IPData, 2, IPData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[IPData, 4, IPData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[IPData, 6, IPData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[IPData, 8, IPData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[IPData, 10, IPData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[IPData, 12, IPData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        wsL.Cells[IPData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[IPData, 2, IPData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[IPData, 4, IPData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[IPData, 6, IPData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[IPData, 8, IPData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[IPData, 10, IPData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[IPData, 12, IPData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        wsL.Cells[IPData, 1].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[IPData, 2, IPData, 3].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[IPData, 4, IPData, 5].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[IPData, 6, IPData, 7].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[IPData, 8, IPData, 9].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[IPData, 10, IPData, 11].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[IPData, 12, IPData, 13].Style.Font.Name = "Arial Narrow";


                        wsL.Cells[IPData, 1].Style.Font.Size = 12;
                        wsL.Cells[IPData, 2, IPData, 3].Style.Font.Size = 12;
                        wsL.Cells[IPData, 4, IPData, 5].Style.Font.Size = 12;
                        wsL.Cells[IPData, 6, IPData, 7].Style.Font.Size = 12;
                        wsL.Cells[IPData, 8, IPData, 9].Style.Font.Size = 12;
                        wsL.Cells[IPData, 10, IPData, 11].Style.Font.Size = 12;
                        wsL.Cells[IPData, 12, IPData, 13].Style.Font.Size = 12;
                        #endregion

                        wsL.Cells[IPData, 1].Value = IPLand.CaseReportID;
                        wsL.Cells[IPData, 2, IPData, 3].Value = IPLand.FamilyName + ", " + IPLand.GivenName + " " + IPLand.MaidenName;
                        wsL.Cells[IPData, 4, IPData, 5].Value = IPLand.DateReported.ToString();
                        wsL.Cells[IPData, 6, IPData, 7].Value = IPLand.CaseLocation;
                        wsL.Cells[IPData, 8, IPData, 9].Value = IPLand.XCoordinates + "," + IPLand.YCoordinates;
                        wsL.Cells[IPData, 10, IPData, 11].Value = IPLand.UpdatedStatusDate.ToString();
                        wsL.Cells[IPData, 12, IPData, 13].Value = IPLand.Handler_FamilyName + ", " + IPLand.Handler_GivenName + " " + IPLand.Handler_MaidenName;

                        IPData++;
                    }
                    
                    int CompletedHeader = wsL.Dimension.End.Row + 1;
                    #region Completed

                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Merge = true;
                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Value = "Completed Reports";
                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.Font.Size = 15;


                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Merge = true;
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Merge = true;
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Merge = true;
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Merge = true;
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Merge = true;
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Merge = true;

                    wsL.Cells[CompletedHeader + 3, 1].Value = "Reference No.";

                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Value = "Username";
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Value = "Date Reported";
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Value = "Location";
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Value = "Coordinates";
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Value = "Date Updated";
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Value = "Handler";

                    //REFERENCE NO
                    wsL.Cells[CompletedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 1].Style.Font.Size = 12;

                    //USERNAME
                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.Font.Size = 12;

                    //DATE REPORTED
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.Font.Size = 12;

                    //LOCATION
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.Font.Size = 12;

                    //COORDINATES
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.Font.Size = 12;

                    //DATE UPDATED
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.Font.Size = 12;

                    //HANDLER
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.Font.Size = 12;

                    int CData = wsL.Dimension.End.Row + 1;
                    #endregion
                    //COMPLETED DATA SECTION
                    foreach (var CLand in _IDataProvider.GetCaseReport(5).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == loopingMonth).ToList())
                    {
                        #region DesignSubmit


                        if (CLand.EnvironmentalConcernID == 1)
                        {
                            wsL.Cells[CData, 1, CData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[CData, 1, CData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                        }
                        else
                        {
                            wsL.Cells[CData, 1, CData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            wsL.Cells[CData, 1, CData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                        }

                        wsL.Cells[CData, 1].Merge = true;
                        wsL.Cells[CData, 2, CData, 3].Merge = true;
                        wsL.Cells[CData, 4, CData, 5].Merge = true;
                        wsL.Cells[CData, 6, CData, 7].Merge = true;
                        wsL.Cells[CData, 8, CData, 9].Merge = true;
                        wsL.Cells[CData, 10, CData, 11].Merge = true;
                        wsL.Cells[CData, 12, CData, 13].Merge = true;

                        wsL.Cells[CData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[CData, 2, CData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[CData, 4, CData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[CData, 6, CData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[CData, 8, CData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[CData, 10, CData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        wsL.Cells[CData, 12, CData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        wsL.Cells[CData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[CData, 2, CData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[CData, 4, CData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[CData, 6, CData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[CData, 8, CData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[CData, 10, CData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        wsL.Cells[CData, 12, CData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        wsL.Cells[CData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[CData, 2, CData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[CData, 4, CData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[CData, 6, CData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[CData, 8, CData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[CData, 10, CData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        wsL.Cells[CData, 12, CData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                        wsL.Cells[CData, 1].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[CData, 2, CData, 3].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[CData, 4, CData, 5].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[CData, 6, CData, 7].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[CData, 8, CData, 9].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[CData, 10, CData, 11].Style.Font.Name = "Arial Narrow";
                        wsL.Cells[CData, 12, CData, 13].Style.Font.Name = "Arial Narrow";


                        wsL.Cells[CData, 1].Style.Font.Size = 12;
                        wsL.Cells[CData, 2, CData, 3].Style.Font.Size = 12;
                        wsL.Cells[CData, 4, CData, 5].Style.Font.Size = 12;
                        wsL.Cells[CData, 6, CData, 7].Style.Font.Size = 12;
                        wsL.Cells[CData, 8, CData, 9].Style.Font.Size = 12;
                        wsL.Cells[CData, 10, CData, 11].Style.Font.Size = 12;
                        wsL.Cells[CData, 12, CData, 13].Style.Font.Size = 12;
                        #endregion

                        wsL.Cells[CData, 1].Value = CLand.CaseReportID;
                        wsL.Cells[CData, 2, CData, 3].Value = CLand.FamilyName + ", " + CLand.GivenName + " " + CLand.MaidenName;
                        wsL.Cells[CData, 4, CData, 5].Value = CLand.DateReported.ToString();
                        wsL.Cells[CData, 6, CData, 7].Value = CLand.CaseLocation;
                        wsL.Cells[CData, 8, CData, 9].Value = CLand.XCoordinates + "," + CLand.YCoordinates;
                        wsL.Cells[CData, 10, CData, 11].Value = CLand.UpdatedStatusDate.ToString();
                        wsL.Cells[CData, 12, CData, 13].Value = CLand.Handler_FamilyName + ", " + CLand.Handler_GivenName + " " + CLand.Handler_MaidenName;

                        CData++;
                    }
                }



                #endregion


                //DOWNLOAD
                Session["DownloadExcel_FileManager"] = excelPackage.GetAsByteArray();
                if (Session["DownloadExcel_FileManager"] != null)
                {
                    byte[] data = Session["DownloadExcel_FileManager"] as byte[];
                    return File(data, "application/octet-stream", "FileManager.xlsx");
                }
                else
                {
                    return new EmptyResult();
                }

            }

        }

        public ActionResult GenerateMonthReport(int year, int month)
        {
            var memoryStream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DateTimeFormatInfo mfi = new DateTimeFormatInfo();
            using (var excelPackage = new ExcelPackage(memoryStream))
            {

                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "VDWWD";
                excelPackage.Workbook.Properties.Title = "Title of Document";
                excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

              
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add(mfi.GetMonthName(month)+" "+year+" Report");
                #region 
                try
                {

                    int height = 230;
                    int width = 230;

                    string DENRLogo = @"C:\Users\pc\Desktop\DENRLogo.png";
                    Image DENRLOGO = Image.FromFile(DENRLogo);
                    ExcelPicture picDENR = ws.Drawings.AddPicture("DENR", DENRLOGO);
                    picDENR.SetPosition(0, 0, 0, 0);//5,3
                    picDENR.SetSize(height, width);

                    string PHLogo = @"C:\Users\pc\Desktop\PHLogo.png";
                    Image PHILIPPINELOGO = Image.FromFile(PHLogo);
                    ExcelPicture picPH = ws.Drawings.AddPicture("PH", PHILIPPINELOGO);
                    picPH.SetPosition(0, 0, 10, 0);//5,3
                    picPH.SetSize(height, width);


                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw;
                }
                //Resize the columns A-M
                ws.Column(1).Width = 24;
                for (int startColumns = 2; startColumns <= 13; startColumns++)
                {
                    ws.Column(startColumns).Width = 12;
                }

                //Resize the row 1-7
                ws.Row(1).Height = 40;
                ws.Row(2).Height = 35;
                ws.Row(3).Height = 25;
                ws.Row(4).Height = 25;
                ws.Row(5).Height = 20;
                ws.Row(6).Height = 20;

                //Merging
                for (int mergeHeader = 1; mergeHeader <= 7; mergeHeader++)
                {
                    ws.Cells[mergeHeader, 1, mergeHeader, 13].Merge = true;
                }

                //Adding Descriptive Values
                ws.Cells[1, 1, 1, 13].Value = "Republic of the Philippines";
                ws.Cells[2, 1, 2, 13].Value = "Department of Environment and National Resources";
                ws.Cells[3, 1, 3, 13].Value = "Environmental Management Bureau";
                ws.Cells[4, 1, 4, 13].Value = "Environmental Monitoring and Enforcement Division";
                ws.Cells[5, 1, 5, 13].Value = "National Ecology Center Compound East Avenue Diliman Quezon City";
                ws.Cells[6, 1, 6, 13].Value = "Tel/Email: 8931-3506 | 8931-2684 | RecordsNCR@emb.gov.ph";

                //Changing Font Size
                ws.Cells[1, 1, 1, 13].Style.Font.Size = 30;
                ws.Cells[2, 1, 2, 13].Style.Font.Size = 25;
                ws.Cells[3, 1, 3, 13].Style.Font.Size = 20;
                ws.Cells[4, 1, 4, 13].Style.Font.Size = 20;
                ws.Cells[5, 1, 5, 13].Style.Font.Size = 16;
                ws.Cells[6, 1, 6, 13].Style.Font.Size = 16;

                //Change the Font Style of the Header
                ws.Cells[1, 1, 6, 13].Style.Font.Name = "Times New Roman";

                //Align to center the header
                ws.Cells[1, 1, 6, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[1, 1, 6, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                #endregion
                #region MainData


                int marginMake = ws.Dimension.End.Row + 2;

                ws.Cells[marginMake, 1, marginMake + 1, 13].Merge = true;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Font.Name = "Arial Narrow";
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Font.Size = 15;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.Font.Bold = true;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[marginMake, 1, marginMake + 1, 13].Value = year + " Summarized Report per City / Municipality";

                int subDataHeaderStart = ws.Dimension.End.Row + 2;




                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Merge = true;
                ws.Cells[subDataHeaderStart, 2, subDataHeaderStart, 4].Merge = true;
                ws.Cells[subDataHeaderStart, 5, subDataHeaderStart, 7].Merge = true;
                ws.Cells[subDataHeaderStart, 8, subDataHeaderStart, 10].Merge = true;
                ws.Cells[subDataHeaderStart, 11, subDataHeaderStart, 13].Merge = true;

                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[subDataHeaderStart, 2, subDataHeaderStart, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[subDataHeaderStart, 5, subDataHeaderStart, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[subDataHeaderStart, 8, subDataHeaderStart, 10].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[subDataHeaderStart, 11, subDataHeaderStart, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Value = mfi.GetMonthName(month).ToString() + " Breakdown City/Municipality";
                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 1].Style.WrapText = true;

                ws.Cells[subDataHeaderStart, 2, subDataHeaderStart, 4].Value = "Submitted";
                ws.Cells[subDataHeaderStart, 5, subDataHeaderStart, 7].Value = "Rejected";
                ws.Cells[subDataHeaderStart, 8, subDataHeaderStart, 10].Value = "Accepted";
                ws.Cells[subDataHeaderStart, 11, subDataHeaderStart, 13].Value = "Completed";

                for (int L = 2; L <= 13; L += 3)
                {
                    ws.Cells[subDataHeaderStart + 1, L].Value = "L";
                    ws.Cells[subDataHeaderStart + 1, L].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[subDataHeaderStart + 1, L].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    ws.Cells[subDataHeaderStart + 1, L + 1].Value = "W";
                    ws.Cells[subDataHeaderStart + 1, L + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[subDataHeaderStart + 1, L + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                    ws.Cells[subDataHeaderStart + 1, L + 2].Value = "TOTAL";
                    ws.Cells[subDataHeaderStart + 1, L + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[subDataHeaderStart + 1, L + 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218));

                    ws.Cells[subDataHeaderStart + 1, L].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 1, L + 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 1, L + 2].Style.Border.BorderAround(ExcelBorderStyle.Double);
                }

                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Font.Name = "Arial Narrow";
                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Font.Size = 12;
                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                ws.Cells[subDataHeaderStart, 1, subDataHeaderStart + 1, 13].Style.Font.Bold = true;

                foreach (var getData in _IDataProvider.GetAreaDetailsPerMonthYear(month, year))
                {
                    for (int L = 2; L <= 13; L += 3)
                    {
                        ws.Cells[subDataHeaderStart + 2, L].Value = "L";
                        ws.Cells[subDataHeaderStart + 2, L].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[subDataHeaderStart + 2, L].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));

                        ws.Cells[subDataHeaderStart + 2, L + 1].Value = "W";
                        ws.Cells[subDataHeaderStart + 2, L + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[subDataHeaderStart + 2, L + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));

                        ws.Cells[subDataHeaderStart + 2, L + 2].Value = "TOTAL";
                        ws.Cells[subDataHeaderStart + 2, L + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[subDataHeaderStart + 2, L + 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(226, 239, 218));
                    }

                    ws.Cells[subDataHeaderStart + 2, 1].Value = getData.CaseLocation;
                    ws.Cells[subDataHeaderStart + 2, 2].Value = getData.L_Submitted;
                    ws.Cells[subDataHeaderStart + 2, 3].Value = getData.W_Submitted;
                    ws.Cells[subDataHeaderStart + 2, 4].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 2] + ":" + ws.Cells[subDataHeaderStart + 2, 3] + ")";

                    ws.Cells[subDataHeaderStart + 2, 5].Value = getData.L_Rejected;
                    ws.Cells[subDataHeaderStart + 2, 6].Value = getData.W_Rejected;
                    ws.Cells[subDataHeaderStart + 2, 7].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 5] + ":" + ws.Cells[subDataHeaderStart + 2, 6] + ")";

                    ws.Cells[subDataHeaderStart + 2, 8].Value = getData.L_Accepted;
                    ws.Cells[subDataHeaderStart + 2, 9].Value = getData.W_Accepted;
                    ws.Cells[subDataHeaderStart + 2, 10].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 8] + ":" + ws.Cells[subDataHeaderStart + 2, 9] + ")";

                    ws.Cells[subDataHeaderStart + 2, 11].Value = getData.L_Completed;
                    ws.Cells[subDataHeaderStart + 2, 12].Value = getData.W_Completed;
                    ws.Cells[subDataHeaderStart + 2, 13].Formula = "SUM(" + ws.Cells[subDataHeaderStart + 2, 11] + ":" + ws.Cells[subDataHeaderStart + 2, 12] + ")";
                    ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.Font.Name = "Arial Narrow";
                    ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.Font.Size = 12;
                    ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[subDataHeaderStart + 2, 1, subDataHeaderStart + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    ws.Cells[subDataHeaderStart + 2, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    ws.Cells[subDataHeaderStart + 2, 2].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    ws.Cells[subDataHeaderStart + 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 6].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    ws.Cells[subDataHeaderStart + 2, 8].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 10].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    ws.Cells[subDataHeaderStart + 2, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 12].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    ws.Cells[subDataHeaderStart + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    subDataHeaderStart++;
                }
                #endregion

                #region GenData
                ExcelWorksheet wsL = excelPackage.Workbook.Worksheets.Add(mfi.GetMonthName(month) + " " + year + " Land Report");

                try
                {
                    int height = 230;
                    int width = 230;

                    string DENRLogo = @"C:\Users\pc\Desktop\DENRLogo.png";
                    Image DENRLOGO = Image.FromFile(DENRLogo);
                    ExcelPicture picDENR = wsL.Drawings.AddPicture("DENR", DENRLOGO);
                    picDENR.SetPosition(0, 0, 0, 0);//5,3
                    picDENR.SetSize(height, width);

                    string PHLogo = @"C:\Users\pc\Desktop\PHLogo.png";
                    Image PHILIPPINELOGO = Image.FromFile(PHLogo);
                    ExcelPicture picPH = wsL.Drawings.AddPicture("PH", PHILIPPINELOGO);
                    picPH.SetPosition(0, 0, 10, 0);//5,3
                    picPH.SetSize(height, width);
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw;
                }


                //Resize the columns A-M
                wsL.Column(1).Width = 24;
                for (int startColumns = 2; startColumns <= 13; startColumns++)
                {
                    wsL.Column(startColumns).Width = 12;
                }

                //Resize the row 1-7
                wsL.Row(1).Height = 40;
                wsL.Row(2).Height = 35;
                wsL.Row(3).Height = 25;
                wsL.Row(4).Height = 25;
                wsL.Row(5).Height = 20;
                wsL.Row(6).Height = 20;

                //Merging
                for (int mergeHeader = 1; mergeHeader <= 7; mergeHeader++)
                {
                    wsL.Cells[mergeHeader, 1, mergeHeader, 13].Merge = true;
                }

                //Adding Descriptive Values
                wsL.Cells[1, 1, 1, 13].Value = "Republic of the Philippines";
                wsL.Cells[2, 1, 2, 13].Value = "Department of Environment and National Resources";
                wsL.Cells[3, 1, 3, 13].Value = "Environmental Management Bureau";
                wsL.Cells[4, 1, 4, 13].Value = "Environmental Monitoring and Enforcement Division";
                wsL.Cells[5, 1, 5, 13].Value = "National Ecology Center Compound East Avenue Diliman Quezon City";
                wsL.Cells[6, 1, 6, 13].Value = "Tel/Email: 8931-3506 | 8931-2684 | RecordsNCR@emb.gov.ph";

                //Changing Font Size
                wsL.Cells[1, 1, 1, 13].Style.Font.Size = 30;
                wsL.Cells[2, 1, 2, 13].Style.Font.Size = 25;
                wsL.Cells[3, 1, 3, 13].Style.Font.Size = 20;
                wsL.Cells[4, 1, 4, 13].Style.Font.Size = 20;
                wsL.Cells[5, 1, 5, 13].Style.Font.Size = 16;
                wsL.Cells[6, 1, 6, 13].Style.Font.Size = 16;

                //Change the Font Style of the Header
                wsL.Cells[1, 1, 6, 13].Style.Font.Name = "Times New Roman";

                //Align to center the header
                wsL.Cells[1, 1, 6, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[1, 1, 6, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                wsL.Cells[8, 1, 9, 13].Merge = true;
                wsL.Cells[8, 1, 9, 13].Value = mfi.GetMonthName(month) + " " + year + " Report";
                wsL.Cells[8, 1, 9, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[8, 1, 9, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[8, 1, 9, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[8, 1, 9, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[8, 1, 9, 13].Style.Font.Size = 15;

                int SubmittedHeader = wsL.Dimension.End.Row + 1;
                #region Submit
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Merge = true;
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Value = "Submitted Reports";
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 1, 1, SubmittedHeader + 2, 13].Style.Font.Size = 15;


                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Merge = true;
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Merge = true;
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Merge = true;
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Merge = true;
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Merge = true;
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Merge = true;

                wsL.Cells[SubmittedHeader + 3, 1].Value = "Reference No.";

                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Value = "Username";
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Value = "Date Reported";
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Value = "Location";
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Value = "Coordinates";
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Value = "Date Updated";
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Value = "Handler";

                //REFERENCE NO
                wsL.Cells[SubmittedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 1].Style.Font.Size = 12;

                //USERNAME
                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 2, SubmittedHeader + 3, 3].Style.Font.Size = 12;

                //DATE REPORTED
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 4, SubmittedHeader + 3, 5].Style.Font.Size = 12;

                //LOCATION
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 6, SubmittedHeader + 3, 7].Style.Font.Size = 12;

                //COORDINATES
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 8, SubmittedHeader + 3, 9].Style.Font.Size = 12;

                //DATE UPDATED
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 10, SubmittedHeader + 3, 11].Style.Font.Size = 12;

                //HANDLER
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[SubmittedHeader + 3, 12, SubmittedHeader + 3, 13].Style.Font.Size = 12;

                int SData = wsL.Dimension.End.Row + 1;
                #endregion
                //SUBMITTED DATA SECTION
                foreach (var SLandSubmit in _IDataProvider.GetCaseReport(1).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == month).OrderBy(o => o.EnvironmentalConcernID).ToList())
                {
                    if (SLandSubmit.EnvironmentalConcernID == 1)
                    {
                        wsL.Cells[SData, 1, SData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[SData, 1, SData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    }
                    else
                    {
                        wsL.Cells[SData, 1, SData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[SData, 1, SData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                    }

                    #region DesignSubmit
                    wsL.Cells[SData, 1].Merge = true;
                    wsL.Cells[SData, 2, SData, 3].Merge = true;
                    wsL.Cells[SData, 4, SData, 5].Merge = true;
                    wsL.Cells[SData, 6, SData, 7].Merge = true;
                    wsL.Cells[SData, 8, SData, 9].Merge = true;
                    wsL.Cells[SData, 10, SData, 11].Merge = true;
                    wsL.Cells[SData, 12, SData, 13].Merge = true;

                    wsL.Cells[SData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SData, 2, SData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SData, 4, SData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SData, 6, SData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SData, 8, SData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SData, 10, SData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[SData, 12, SData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    wsL.Cells[SData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SData, 2, SData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SData, 4, SData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SData, 6, SData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SData, 8, SData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SData, 10, SData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[SData, 12, SData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    wsL.Cells[SData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SData, 2, SData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SData, 4, SData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SData, 6, SData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SData, 8, SData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SData, 10, SData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[SData, 12, SData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    wsL.Cells[SData, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SData, 2, SData, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SData, 4, SData, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SData, 6, SData, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SData, 8, SData, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SData, 10, SData, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[SData, 12, SData, 13].Style.Font.Name = "Arial Narrow";


                    wsL.Cells[SData, 1].Style.Font.Size = 12;
                    wsL.Cells[SData, 2, SData, 3].Style.Font.Size = 12;
                    wsL.Cells[SData, 4, SData, 5].Style.Font.Size = 12;
                    wsL.Cells[SData, 6, SData, 7].Style.Font.Size = 12;
                    wsL.Cells[SData, 8, SData, 9].Style.Font.Size = 12;
                    wsL.Cells[SData, 10, SData, 11].Style.Font.Size = 12;
                    wsL.Cells[SData, 12, SData, 13].Style.Font.Size = 12;
                    #endregion

                    wsL.Cells[SData, 1].Value = SLandSubmit.CaseReportID;
                    wsL.Cells[SData, 2, SData, 3].Value = SLandSubmit.FamilyName + ", " + SLandSubmit.GivenName + " " + SLandSubmit.MaidenName;
                    wsL.Cells[SData, 4, SData, 5].Value = SLandSubmit.DateReported.ToString();
                    wsL.Cells[SData, 6, SData, 7].Value = SLandSubmit.CaseLocation;
                    wsL.Cells[SData, 8, SData, 9].Value = SLandSubmit.XCoordinates + "," + SLandSubmit.YCoordinates;
                    wsL.Cells[SData, 10, SData, 11].Value = SLandSubmit.UpdatedStatusDate.ToString();
                    wsL.Cells[SData, 12, SData, 13].Value = SLandSubmit.Handler_FamilyName + ", " + SLandSubmit.Handler_GivenName + " " + SLandSubmit.Handler_MaidenName;

                    SData++;
                }

                int AcceptedHeader = wsL.Dimension.End.Row + 1;
                #region Accept
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Merge = true;
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Value = "Accepted Reports";
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 1, 1, AcceptedHeader + 2, 13].Style.Font.Size = 15;


                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Merge = true;
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Merge = true;
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Merge = true;
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Merge = true;
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Merge = true;
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Merge = true;

                wsL.Cells[AcceptedHeader + 3, 1].Value = "Reference No.";

                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Value = "Username";
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Value = "Date Reported";
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Value = "Location";
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Value = "Coordinates";
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Value = "Date Updated";
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Value = "Handler";

                //REFERENCE NO
                wsL.Cells[AcceptedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 1].Style.Font.Size = 12;

                //USERNAME
                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 2, AcceptedHeader + 3, 3].Style.Font.Size = 12;

                //DATE REPORTED
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 4, AcceptedHeader + 3, 5].Style.Font.Size = 12;

                //LOCATION
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 6, AcceptedHeader + 3, 7].Style.Font.Size = 12;

                //COORDINATES
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 8, AcceptedHeader + 3, 9].Style.Font.Size = 12;

                //DATE UPDATED
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 10, AcceptedHeader + 3, 11].Style.Font.Size = 12;

                //HANDLER
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[AcceptedHeader + 3, 12, AcceptedHeader + 3, 13].Style.Font.Size = 12;

                int AData = wsL.Dimension.End.Row + 1;

                #endregion
                //ACCEPTED DATA SECTION
                foreach (var SLandAccept in _IDataProvider.GetCaseReport(3).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == month).ToList())
                {
                    if (SLandAccept.EnvironmentalConcernID == 1)
                    {
                        wsL.Cells[AData, 1, AData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[AData, 1, AData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    }
                    else
                    {
                        wsL.Cells[AData, 1, AData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[AData, 1, AData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                    }

                    #region DesignSubmit
                    wsL.Cells[AData, 1].Merge = true;
                    wsL.Cells[AData, 2, AData, 3].Merge = true;
                    wsL.Cells[AData, 4, AData, 5].Merge = true;
                    wsL.Cells[AData, 6, AData, 7].Merge = true;
                    wsL.Cells[AData, 8, AData, 9].Merge = true;
                    wsL.Cells[AData, 10, AData, 11].Merge = true;
                    wsL.Cells[AData, 12, AData, 13].Merge = true;

                    wsL.Cells[AData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AData, 2, AData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AData, 4, AData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AData, 6, AData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AData, 8, AData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AData, 10, AData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[AData, 12, AData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    wsL.Cells[AData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AData, 2, AData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AData, 4, AData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AData, 6, AData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AData, 8, AData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AData, 10, AData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[AData, 12, AData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    wsL.Cells[AData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AData, 2, AData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AData, 4, AData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AData, 6, AData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AData, 8, AData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AData, 10, AData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[AData, 12, AData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    wsL.Cells[AData, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AData, 2, AData, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AData, 4, AData, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AData, 6, AData, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AData, 8, AData, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AData, 10, AData, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[AData, 12, AData, 13].Style.Font.Name = "Arial Narrow";


                    wsL.Cells[AData, 1].Style.Font.Size = 12;
                    wsL.Cells[AData, 2, AData, 3].Style.Font.Size = 12;
                    wsL.Cells[AData, 4, AData, 5].Style.Font.Size = 12;
                    wsL.Cells[AData, 6, AData, 7].Style.Font.Size = 12;
                    wsL.Cells[AData, 8, AData, 9].Style.Font.Size = 12;
                    wsL.Cells[AData, 10, AData, 11].Style.Font.Size = 12;
                    wsL.Cells[AData, 12, AData, 13].Style.Font.Size = 12;
                    #endregion

                    wsL.Cells[AData, 1].Value = SLandAccept.CaseReportID;
                    wsL.Cells[AData, 2, AData, 3].Value = SLandAccept.FamilyName + ", " + SLandAccept.GivenName + " " + SLandAccept.MaidenName;
                    wsL.Cells[AData, 4, AData, 5].Value = SLandAccept.DateReported.ToString();
                    wsL.Cells[AData, 6, AData, 7].Value = SLandAccept.CaseLocation;
                    wsL.Cells[AData, 8, AData, 9].Value = SLandAccept.XCoordinates + "," + SLandAccept.YCoordinates;
                    wsL.Cells[AData, 10, AData, 11].Value = SLandAccept.UpdatedStatusDate.ToString();
                    wsL.Cells[AData, 12, AData, 13].Value = SLandAccept.Handler_FamilyName + ", " + SLandAccept.Handler_GivenName + " " + SLandAccept.Handler_MaidenName;

                    AData++;
                }

                int RejectedHeader = wsL.Dimension.End.Row + 1;
                #region Reject
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Merge = true;
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Value = "Rejected Reports";
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 1, 1, RejectedHeader + 2, 13].Style.Font.Size = 15;


                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Merge = true;
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Merge = true;
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Merge = true;
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Merge = true;
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Merge = true;
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Merge = true;

                wsL.Cells[RejectedHeader + 3, 1].Value = "Reference No.";

                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Value = "Username";
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Value = "Date Reported";
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Value = "Location";
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Value = "Coordinates";
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Value = "Date Updated";
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Value = "Handler";

                //REFERENCE NO
                wsL.Cells[RejectedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 1].Style.Font.Size = 12;

                //USERNAME
                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 2, RejectedHeader + 3, 3].Style.Font.Size = 12;

                //DATE REPORTED
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 4, RejectedHeader + 3, 5].Style.Font.Size = 12;

                //LOCATION
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 6, RejectedHeader + 3, 7].Style.Font.Size = 12;

                //COORDINATES
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 8, RejectedHeader + 3, 9].Style.Font.Size = 12;

                //DATE UPDATED
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 10, RejectedHeader + 3, 11].Style.Font.Size = 12;

                //HANDLER
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[RejectedHeader + 3, 12, RejectedHeader + 3, 13].Style.Font.Size = 12;

                int RData = wsL.Dimension.End.Row + 1;
                #endregion
                //REJECTED DATA SECTION
                foreach (var LLandReject in _IDataProvider.GetCaseReport(2).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == month).ToList())
                {
                    #region DesignSubmit
                    wsL.Cells[RData, 1].Merge = true;
                    wsL.Cells[RData, 2, RData, 3].Merge = true;
                    wsL.Cells[RData, 4, RData, 5].Merge = true;
                    wsL.Cells[RData, 6, RData, 7].Merge = true;
                    wsL.Cells[RData, 8, RData, 9].Merge = true;
                    wsL.Cells[RData, 10, RData, 11].Merge = true;
                    wsL.Cells[RData, 12, RData, 13].Merge = true;

                    wsL.Cells[RData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RData, 2, RData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RData, 4, RData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RData, 6, RData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RData, 8, RData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RData, 10, RData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[RData, 12, RData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    wsL.Cells[RData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RData, 2, RData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RData, 4, RData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RData, 6, RData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RData, 8, RData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RData, 10, RData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[RData, 12, RData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    wsL.Cells[RData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RData, 2, RData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RData, 4, RData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RData, 6, RData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RData, 8, RData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RData, 10, RData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[RData, 12, RData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    wsL.Cells[RData, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RData, 2, RData, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RData, 4, RData, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RData, 6, RData, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RData, 8, RData, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RData, 10, RData, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[RData, 12, RData, 13].Style.Font.Name = "Arial Narrow";


                    wsL.Cells[RData, 1].Style.Font.Size = 12;
                    wsL.Cells[RData, 2, RData, 3].Style.Font.Size = 12;
                    wsL.Cells[RData, 4, RData, 5].Style.Font.Size = 12;
                    wsL.Cells[RData, 6, RData, 7].Style.Font.Size = 12;
                    wsL.Cells[RData, 8, RData, 9].Style.Font.Size = 12;
                    wsL.Cells[RData, 10, RData, 11].Style.Font.Size = 12;
                    wsL.Cells[RData, 12, RData, 13].Style.Font.Size = 12;
                    #endregion

                    if (LLandReject.EnvironmentalConcernID == 1)
                    {
                        wsL.Cells[RData, 1, RData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[RData, 1, RData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    }
                    else
                    {
                        wsL.Cells[RData, 1, RData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[RData, 1, RData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                    }

                    wsL.Cells[RData, 1].Value = LLandReject.CaseReportID;
                    wsL.Cells[RData, 2, RData, 3].Value = LLandReject.FamilyName + ", " + LLandReject.GivenName + " " + LLandReject.MaidenName;
                    wsL.Cells[RData, 4, RData, 5].Value = LLandReject.DateReported.ToString();
                    wsL.Cells[RData, 6, RData, 7].Value = LLandReject.CaseLocation;
                    wsL.Cells[RData, 8, RData, 9].Value = LLandReject.XCoordinates + "," + LLandReject.YCoordinates;
                    wsL.Cells[RData, 10, RData, 11].Value = LLandReject.UpdatedStatusDate.ToString();
                    wsL.Cells[RData, 12, RData, 13].Value = LLandReject.Handler_FamilyName + ", " + LLandReject.Handler_GivenName + " " + LLandReject.Handler_MaidenName;

                    RData++;
                }

                int InProgressHeader = wsL.Dimension.End.Row + 1;
                #region InProgress
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Merge = true;
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Value = "In Progress Reports";
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 1, 1, InProgressHeader + 2, 13].Style.Font.Size = 15;


                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Merge = true;
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Merge = true;
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Merge = true;
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Merge = true;
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Merge = true;
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Merge = true;

                wsL.Cells[InProgressHeader + 3, 1].Value = "Reference No.";

                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Value = "Username";
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Value = "Date Reported";
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Value = "Location";
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Value = "Coordinates";
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Value = "Date Updated";
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Value = "Handler";

                //REFERENCE NO
                wsL.Cells[InProgressHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 1].Style.Font.Size = 12;

                //USERNAME
                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 2, InProgressHeader + 3, 3].Style.Font.Size = 12;

                //DATE REPORTED
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 4, InProgressHeader + 3, 5].Style.Font.Size = 12;

                //LOCATION
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 6, InProgressHeader + 3, 7].Style.Font.Size = 12;

                //COORDINATES
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 8, InProgressHeader + 3, 9].Style.Font.Size = 12;

                //DATE UPDATED
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 10, InProgressHeader + 3, 11].Style.Font.Size = 12;

                //HANDLER
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[InProgressHeader + 3, 12, InProgressHeader + 3, 13].Style.Font.Size = 12;

                int IPData = wsL.Dimension.End.Row + 1;
                #endregion
                //  IN PROGRESS DATA SECTION
                foreach (var IPLand in _IDataProvider.GetCaseReport(4).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == month).ToList())
                {
                    #region DesignSubmit

                    if (IPLand.EnvironmentalConcernID == 1)
                    {
                        wsL.Cells[IPData, 1, IPData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[IPData, 1, IPData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    }
                    else
                    {
                        wsL.Cells[IPData, 1, IPData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[IPData, 1, IPData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                    }


                    wsL.Cells[IPData, 1].Merge = true;
                    wsL.Cells[IPData, 2, IPData, 3].Merge = true;
                    wsL.Cells[IPData, 4, IPData, 5].Merge = true;
                    wsL.Cells[IPData, 6, IPData, 7].Merge = true;
                    wsL.Cells[IPData, 8, IPData, 9].Merge = true;
                    wsL.Cells[IPData, 10, IPData, 11].Merge = true;
                    wsL.Cells[IPData, 12, IPData, 13].Merge = true;

                    wsL.Cells[IPData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[IPData, 2, IPData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[IPData, 4, IPData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[IPData, 6, IPData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[IPData, 8, IPData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[IPData, 10, IPData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[IPData, 12, IPData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    wsL.Cells[IPData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[IPData, 2, IPData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[IPData, 4, IPData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[IPData, 6, IPData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[IPData, 8, IPData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[IPData, 10, IPData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[IPData, 12, IPData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    wsL.Cells[IPData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[IPData, 2, IPData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[IPData, 4, IPData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[IPData, 6, IPData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[IPData, 8, IPData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[IPData, 10, IPData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[IPData, 12, IPData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    wsL.Cells[IPData, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[IPData, 2, IPData, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[IPData, 4, IPData, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[IPData, 6, IPData, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[IPData, 8, IPData, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[IPData, 10, IPData, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[IPData, 12, IPData, 13].Style.Font.Name = "Arial Narrow";


                    wsL.Cells[IPData, 1].Style.Font.Size = 12;
                    wsL.Cells[IPData, 2, IPData, 3].Style.Font.Size = 12;
                    wsL.Cells[IPData, 4, IPData, 5].Style.Font.Size = 12;
                    wsL.Cells[IPData, 6, IPData, 7].Style.Font.Size = 12;
                    wsL.Cells[IPData, 8, IPData, 9].Style.Font.Size = 12;
                    wsL.Cells[IPData, 10, IPData, 11].Style.Font.Size = 12;
                    wsL.Cells[IPData, 12, IPData, 13].Style.Font.Size = 12;
                    #endregion

                    wsL.Cells[IPData, 1].Value = IPLand.CaseReportID;
                    wsL.Cells[IPData, 2, IPData, 3].Value = IPLand.FamilyName + ", " + IPLand.GivenName + " " + IPLand.MaidenName;
                    wsL.Cells[IPData, 4, IPData, 5].Value = IPLand.DateReported.ToString();
                    wsL.Cells[IPData, 6, IPData, 7].Value = IPLand.CaseLocation;
                    wsL.Cells[IPData, 8, IPData, 9].Value = IPLand.XCoordinates + "," + IPLand.YCoordinates;
                    wsL.Cells[IPData, 10, IPData, 11].Value = IPLand.UpdatedStatusDate.ToString();
                    wsL.Cells[IPData, 12, IPData, 13].Value = IPLand.Handler_FamilyName + ", " + IPLand.Handler_GivenName + " " + IPLand.Handler_MaidenName;

                    IPData++;
                }

                int CompletedHeader = wsL.Dimension.End.Row + 1;
                #region Completed

                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Merge = true;
                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Value = "Completed Reports";
                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 1, 1, CompletedHeader + 2, 13].Style.Font.Size = 15;


                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Merge = true;
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Merge = true;
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Merge = true;
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Merge = true;
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Merge = true;
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Merge = true;

                wsL.Cells[CompletedHeader + 3, 1].Value = "Reference No.";

                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Value = "Username";
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Value = "Date Reported";
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Value = "Location";
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Value = "Coordinates";
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Value = "Date Updated";
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Value = "Handler";

                //REFERENCE NO
                wsL.Cells[CompletedHeader + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 1].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 1].Style.Font.Size = 12;

                //USERNAME
                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 2, CompletedHeader + 3, 3].Style.Font.Size = 12;

                //DATE REPORTED
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 4, CompletedHeader + 3, 5].Style.Font.Size = 12;

                //LOCATION
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 6, CompletedHeader + 3, 7].Style.Font.Size = 12;

                //COORDINATES
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 8, CompletedHeader + 3, 9].Style.Font.Size = 12;

                //DATE UPDATED
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 10, CompletedHeader + 3, 11].Style.Font.Size = 12;

                //HANDLER
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.Font.Name = "Arial Narrow";
                wsL.Cells[CompletedHeader + 3, 12, CompletedHeader + 3, 13].Style.Font.Size = 12;

                int CData = wsL.Dimension.End.Row + 1;
                #endregion
                //COMPLETED DATA SECTION
                foreach (var CLand in _IDataProvider.GetCaseReport(5).Where(x => x.UpdatedStatusDate.Year == year && x.UpdatedStatusDate.Month == month).ToList())
                {
                    #region DesignSubmit


                    if (CLand.EnvironmentalConcernID == 1)
                    {
                        wsL.Cells[CData, 1, CData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[CData, 1, CData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 242, 204));
                    }
                    else
                    {
                        wsL.Cells[CData, 1, CData, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        wsL.Cells[CData, 1, CData, 13].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 235, 247));
                    }

                    wsL.Cells[CData, 1].Merge = true;
                    wsL.Cells[CData, 2, CData, 3].Merge = true;
                    wsL.Cells[CData, 4, CData, 5].Merge = true;
                    wsL.Cells[CData, 6, CData, 7].Merge = true;
                    wsL.Cells[CData, 8, CData, 9].Merge = true;
                    wsL.Cells[CData, 10, CData, 11].Merge = true;
                    wsL.Cells[CData, 12, CData, 13].Merge = true;

                    wsL.Cells[CData, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CData, 2, CData, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CData, 4, CData, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CData, 6, CData, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CData, 8, CData, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CData, 10, CData, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsL.Cells[CData, 12, CData, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    wsL.Cells[CData, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CData, 2, CData, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CData, 4, CData, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CData, 6, CData, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CData, 8, CData, 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CData, 10, CData, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsL.Cells[CData, 12, CData, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    wsL.Cells[CData, 1].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CData, 2, CData, 3].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CData, 4, CData, 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CData, 6, CData, 7].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CData, 8, CData, 9].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CData, 10, CData, 11].Style.Border.BorderAround(ExcelBorderStyle.Double);
                    wsL.Cells[CData, 12, CData, 13].Style.Border.BorderAround(ExcelBorderStyle.Double);

                    wsL.Cells[CData, 1].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CData, 2, CData, 3].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CData, 4, CData, 5].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CData, 6, CData, 7].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CData, 8, CData, 9].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CData, 10, CData, 11].Style.Font.Name = "Arial Narrow";
                    wsL.Cells[CData, 12, CData, 13].Style.Font.Name = "Arial Narrow";


                    wsL.Cells[CData, 1].Style.Font.Size = 12;
                    wsL.Cells[CData, 2, CData, 3].Style.Font.Size = 12;
                    wsL.Cells[CData, 4, CData, 5].Style.Font.Size = 12;
                    wsL.Cells[CData, 6, CData, 7].Style.Font.Size = 12;
                    wsL.Cells[CData, 8, CData, 9].Style.Font.Size = 12;
                    wsL.Cells[CData, 10, CData, 11].Style.Font.Size = 12;
                    wsL.Cells[CData, 12, CData, 13].Style.Font.Size = 12;
                    #endregion

                    wsL.Cells[CData, 1].Value = CLand.CaseReportID;
                    wsL.Cells[CData, 2, CData, 3].Value = CLand.FamilyName + ", " + CLand.GivenName + " " + CLand.MaidenName;
                    wsL.Cells[CData, 4, CData, 5].Value = CLand.DateReported.ToString();
                    wsL.Cells[CData, 6, CData, 7].Value = CLand.CaseLocation;
                    wsL.Cells[CData, 8, CData, 9].Value = CLand.XCoordinates + "," + CLand.YCoordinates;
                    wsL.Cells[CData, 10, CData, 11].Value = CLand.UpdatedStatusDate.ToString();
                    wsL.Cells[CData, 12, CData, 13].Value = CLand.Handler_FamilyName + ", " + CLand.Handler_GivenName + " " + CLand.Handler_MaidenName;

                    CData++;
                }
                #endregion


                //DOWNLOAD
                Session["DownloadExcel_FileManager"] = excelPackage.GetAsByteArray();
                if (Session["DownloadExcel_FileManager"] != null)
                {
                   
                    byte[] data = Session["DownloadExcel_FileManager"] as byte[];
                    return File(data, "application/octet-stream", "FileManager.xlsx");
                }
                else
                {
                    return new EmptyResult();
                }
            }
        }

    }
}
