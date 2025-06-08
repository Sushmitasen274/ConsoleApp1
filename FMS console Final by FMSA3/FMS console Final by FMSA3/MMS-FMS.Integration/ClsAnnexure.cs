using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Data;
using System.Reflection;
using System.Configuration;
using DBLayer;
using System.Threading.Tasks;
using DBLayer;
using System.Net;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MMS_FMS.Integration
{
    public class ClsAnnexure
    {
        private static string host = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPFILEPATH"]);
        private static string username = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPUSER"]);
        private static string password = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPPASSWORD"]);
        private static string port = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPPORT"]);
        private static string strLocalDir = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["LOCALDIRECTORY"]);
        private static string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
        private static FtpWebRequest ftpRequest = null;
        private static FtpWebResponse ftpResponse = null;


        static DataSet dsDetailsAnnexure1 = new DataSet();
        static DataSet dsDetailsAnnexure2 = new DataSet();
        static DataTable dtDetailsAnnexure2A = new DataTable();
        static DataTable dtDetailsAnnexure3 = new DataTable();
        static DataTable dtDetailsAnnexure5 = new DataTable();
        static DataSet dtDetailsAnnexure6 = new DataSet();
        static DataTable dtDetailsAnnexure3Sept = new DataTable();
        static DataTable dtDetailsAnnexure3MF = new DataTable();


        internal void GetAnnexure(DataTable dtWorkorder)
        {
            string sDestinationPath, sXlsxDestinationFile;
            string sMessage = string.Empty;
            FileInfo ExcelCopy;
            int querts = 0;
            int quertss = 0;
            string strPath = string.Empty;
            string strappend = string.Empty;
            string sStoreName = string.Empty;
            bool blStatus = false;
            bool checkValidation = false;
            DataTable dtmonths = new DataTable();
            ClsAnnexure objAnnex = new ClsAnnexure();

            string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;

            string sXLsxFileName = @"D:\ExcelWorkbook\MARCH_FINAL_ANNEXURES.xlsx";
            //string sXLsxFileName = Server.MapPath("~") + "\\ExcelWorkbook\\MARCH_FINAL_ANNEXURES.xlsx";
            string sXLsxFileNames = "MARCH_FINAL_ANNEXURES.xlsx";
            sDestinationPath = CreateNewFolder(Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]), sXLsxFileNames);
            sXlsxDestinationFile = System.IO.Path.Combine(sDestinationPath, sXLsxFileNames);

            ClsFetchannexure objFetch = new ClsFetchannexure();


            FileInfo file = new FileInfo(sXlsxDestinationFile);
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {

                dsDetailsAnnexure1 = new DataSet();
                dsDetailsAnnexure2 = new DataSet();
                dtDetailsAnnexure2A = new DataTable();
                dtDetailsAnnexure3 = new DataTable();
                dtDetailsAnnexure5 = new DataTable();
                dtDetailsAnnexure6 = new DataSet();
                dtDetailsAnnexure3Sept = new DataTable();
                dtDetailsAnnexure3MF = new DataTable();

                List<Task> TaskList = new List<Task>();
                for (int z = 0; z <= 2; z++)
                {
                    var LastTask = new Task<bool>(() => ProcessRecords(file, dtWorkorder, z, querts, quertss));
                    LastTask.Start();
                    TaskList.Add(LastTask);
                    System.Threading.Thread.Sleep(1000);
                }
                Task.WaitAll(TaskList.ToArray());


                dtmonths = objFetch.getmonthname(Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));
                string LocName = objFetch.GetLocationName(Convert.ToInt32(dtWorkorder.Rows[0]["AL_LOCATION"]));
                if (dsDetailsAnnexure1.Tables.Count > 0)
                {
                    ExportAnnxOneDataLatest(file, dsDetailsAnnexure1, dtmonths, LocName);
                }
                if (dsDetailsAnnexure2.Tables.Count > 0)
                {
                    ExportAnnxTwoData(file, dsDetailsAnnexure2, dtmonths, LocName);
                }

                if (dtDetailsAnnexure2A.Rows.Count > 0)
                {
                    ExportAnnx2AData(file, dtDetailsAnnexure2A, dtmonths, LocName);//krithika
                }
                if (dtDetailsAnnexure3.Rows.Count > 0)
                {
                    ExportAnnxThreeeData(file, dtDetailsAnnexure3, dtmonths, LocName);
                    //krithika
                }
                if (dtDetailsAnnexure5.Rows.Count > 0)
                {
                    ExportAnnxFiveData(file, dtDetailsAnnexure5, dtmonths, LocName);
                }
                if (dtDetailsAnnexure6.Tables.Count > 0)
                {
                    ExportAnnxSixData(file, dtDetailsAnnexure6, dtmonths, LocName);//pradeep
                }
                if (dtDetailsAnnexure3Sept.Rows.Count > 0)
                {
                    ExportAnnxThreeeDataUptoSep(file, dtDetailsAnnexure3Sept, dtmonths, LocName);//madan'
                }
                if (dtDetailsAnnexure3MF.Rows.Count > 0)
                {
                    ExportAnnxThreeeDataUptoMf(file, dtDetailsAnnexure3MF, dtmonths, LocName);//madan'
                }






                strPath = "ftp://" + host + "/";

                strappend = Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]) + "/";



                bool blCheck = CreateDirectory(strPath, strappend);
                var sLocalDir = strLocalDir + "\\" + strappend;

                if (!Directory.Exists(sLocalDir))
                {
                    Directory.CreateDirectory(sLocalDir);
                }

                //  var Headerfile = @"" + strLocalDir + "\\" + Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]).Replace('/', '$') + "_" + "annexures_1_to_8" + ".csv";
                var Headerfile = @"" + strLocalDir + "\\" + "annexures_1_to_8" + "_" + Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]).Replace('/', '$') + ".csv";

                if (blCheck == true)
                {

                    //using (var stream = File.Create(Headerfile))
                    //{
                    //    stream.wr(sXlsxDestinationFile);
                    //}
                    // strPath = "ftp://" + host + "/" + Convert.ToString(dsWorkOrder.Tables[0].Rows[i]["WH_LocationCode"])+"-" + sStoreName + "/";
                    strPath = "ftp://" + host + "/" + Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]) + "/";
                    strappend = Convert.ToString(dtWorkorder.Rows[0]["AL_ID"]) + "/";
                    var filename1 = Path.GetFileName(Headerfile);
                    blCheck = CreateDirectory(strPath, strappend);
                    if (blCheck == true)
                    {
                        deleteDirectoryFTP(strPath + strappend, "cescmysore\ftp_fms", "Idea@123");
                        blCheck = CreateDirectory(strPath, strappend);

                        blStatus = uploadToFtp(sXlsxDestinationFile, strPath + strappend);

                        string sSqlll = "UPDATE TBLANNEXURE_LOG SET AL_ENTRY_FLAG=2,AL_PATH='" + (strPath + strappend + filename1) + "',AL_UPDATED_DATE=getdate() WHERE  AL_ID=" + dtWorkorder.Rows[0]["AL_ID"] + " ";
                        DBHelper.DBExecuteNoNQuery(sConString, sSqlll);
                        //}
                    }
                }


            }

        }

        private static bool deleteDirectoryFTP(string directoryPath, string ftpUser, string ftpPassword)
        {
            try
            {
                FtpWebRequest ftpRequest11 = (FtpWebRequest)WebRequest.Create(directoryPath);
                ftpRequest11.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpRequest11.Method = WebRequestMethods.Ftp.ListDirectory;
                FtpWebResponse response11 = (FtpWebResponse)ftpRequest11.GetResponse();
                StreamReader streamReader = new StreamReader(response11.GetResponseStream());
                string line = streamReader.ReadLine();
                FtpWebRequest ftpRequest1 = null;
                FtpWebResponse response1 = null;
                while (!string.IsNullOrEmpty(line))
                {

                    ftpRequest1 = (FtpWebRequest)WebRequest.Create(directoryPath + "/" + line);
                    ftpRequest1.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                    ftpRequest1.Method = WebRequestMethods.Ftp.DeleteFile;
                    response1 = (FtpWebResponse)ftpRequest1.GetResponse();
                    response1.Close();
                    line = streamReader.ReadLine();
                }
                streamReader.Close();
                response11.Close();
                return true;

            }
            catch (WebException ex)
            {
                AppException.WritetoFile(ex.Message, "", System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
                return false;
            }

        }

        private static bool FtpDirectoryExists(string directoryPath, string ftpUser, string ftpPassword)
        {
            bool directoryExists;

            var request = (FtpWebRequest)WebRequest.Create(directoryPath);
            request.Method = WebRequestMethods.Ftp.ListDirectory;
            request.Credentials = new NetworkCredential(ftpUser, ftpPassword);
            try
            {
                using (request.GetResponse())
                {
                    directoryExists = true;
                }


            }
            catch (WebException ex)
            {
                AppException.WritetoFile(ex.Message, "", System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
                directoryExists = false;
            }

            return directoryExists;
        }

        public static bool ProcessRecords(FileInfo file, DataTable dtWorkorder, int i, int querts, int quertss)
        {
            ClsFetchannexure objfetch = new ClsFetchannexure();
            if (i == 0)
            {


                //AppException.LogErrorAnnexure("Before call sp ANNEXURE 1", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                //    DataRow annex1 = dtWorkorder.AsEnumerable().Contains(dtWorkorder.Rows["AL_ANNEXURE_SUBTYPE"].ToString() == "Annx 1");

                //if (dtWorkorder.Rows[0]["AL_ANNEXURE_TYPE"] == "A1")
                // {


                string filter = "AL_ANNEXURE_SUBTYPE LIKE 'Annx 1%'";
                DataView viewannexure1 = new DataView(dtWorkorder);
                DataTable dtannx1 = new DataTable();
                viewannexure1.RowFilter = filter;
                dtannx1 = viewannexure1.ToTable();

                if (dtannx1.Rows.Count > 0)
                {
                    dsDetailsAnnexure1 = objfetch.GetAnnexureDetailOneLatest(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));
                }

                string filter2 = "AL_ANNEXURE_SUBTYPE LIKE '%Annx 2%'";
                DataView viewannexure2 = new DataView(dtWorkorder);
                DataTable dtannx2 = new DataTable();
                viewannexure2.RowFilter = filter2;
                dtannx2 = viewannexure2.ToTable();

                if (dtannx2.Rows.Count > 0)
                {
                    dsDetailsAnnexure2 = objfetch.GetAnnexureDetailsTwo(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));


                    dtDetailsAnnexure2A = objfetch.GetAnnexure2ADetails(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));

                }

                string filter3 = "AL_ANNEXURE_SUBTYPE LIKE '%Annx 3%'";
                DataView viewannexure3 = new DataView(dtWorkorder);
                DataTable dtannx3 = new DataTable();
                viewannexure3.RowFilter = filter3;
                dtannx3 = viewannexure3.ToTable();

                if (dtannx3.Rows.Count > 0)
                {
                    dtDetailsAnnexure3 = objfetch.GetAnnexureDetailsThree(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));

                    dtDetailsAnnexure3Sept = objfetch.GetAnnexureDetailsThreeuptoSep(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));

                    dtDetailsAnnexure3MF = objfetch.GetAnnexureDetailsThreeupMF(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));
                }



                string filter5 = "AL_ANNEXURE_SUBTYPE LIKE '%Annx 5%'";
                DataView viewannexure5 = new DataView(dtWorkorder);
                DataTable dtannx5 = new DataTable();
                viewannexure5.RowFilter = filter5;
                dtannx5 = viewannexure5.ToTable();

                if (dtannx5.Rows.Count > 0)
                {
                    dtDetailsAnnexure5 = objfetch.GetAnnexureDetailsFive(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));

                }



                string filter6 = "AL_ANNEXURE_SUBTYPE LIKE '%Annx 6%'";
                DataView viewannexure6 = new DataView(dtWorkorder);
                DataTable dtannx6 = new DataTable();
                viewannexure5.RowFilter = filter6;
                dtannx6 = viewannexure6.ToTable();

                if (dtannx6.Rows.Count > 0)
                {
                    dtDetailsAnnexure6 = objfetch.GetAnnexure6(Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]));

                }



                //AppException.LogErrorAnnexure("Before call sp  ANNEXURE 2", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));

                //objAnnex. dtDetailsAnnexure2A = objAnnexure.GetAnnexure2ADetails(objAnnexure.LocationCode.ToString(), objAnnexure.frmonth.ToString(), objAnnexure.tomonth.ToString());
                //AppException.LogErrorAnnexure("After call sp ANNEXURE 2A", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));

                //dtDetailsAnnexure3 = objAnnexure.GetAnnexureDetailsThree(objAnnexure.LocationCode, 4, objAnnexure.frmonth, objAnnexure.tomonth);
                //AppException.LogErrorAnnexure("Before call sp  ANNEXURE 3", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));

                //objAnnex. dtDetailsAnnexure5 = objAnnexure.GetAnnexureDetailsFive(objAnnexure.LocationCode, Convert.ToInt32(objAnnexure.frmonth), Convert.ToInt32(objAnnexure.tomonth), 4);
                //AppException.LogErrorAnnexure("After call sp ANNEXURE 5", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));

                //AppException.LogErrorAnnexure("Before call sp ANNEXURE 6", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                //dtDetailsAnnexure6 = objAnnexure.GetAnnexure6(objAnnexure.LocationCode.ToString(), objAnnexure.frmonth, objAnnexure.tomonth);
                //AppException.LogErrorAnnexure("After call sp ANNEXURE 6", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));

                //if (querts != 0)
                //{
                //    AppException.LogErrorAnnexure("Before call sp  sp Annexure 3SEPT", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                //    objAnnex. dtDetailsAnnexure3Sept = objAnnexure.GetAnnexureDetailsThreeuptoSep(objAnnexure.LocationCode, 4, objAnnexure.frmonth, Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 5));
                //    AppException.LogErrorAnnexure("After call sp Annexure 3SEPT", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                //    //  ExportAnnxThreeeDataUptoSep(file, objAnnexure);//madan'
                //}
                //if (quertss != 0)
                //{
                //    AppException.LogErrorAnnexure("Before call sp Annexure 3MF", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                //    objAnnex. dtDetailsAnnexure3MF = objAnnexure.GetAnnexureDetailsThreeupMF(objAnnexure.LocationCode, 4, Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 6), Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 12));

                //    AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                //    ///ExportAnnxThreeeDataUptoMf(file, objAnnexure);//madan'
                //}

                // }
                return true;


            }
            else
            {
                //  AppException.LogErrorAnnexure("Before call sp ANNEXURE 4 ", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));

                string filter4 = "AL_ANNEXURE_SUBTYPE LIKE '%Annx 4%'";
                DataView viewannexure4 = new DataView(dtWorkorder);
                DataTable dtannx4 = new DataTable();
                viewannexure4.RowFilter = filter4;
                dtannx4 = viewannexure4.ToTable();

                if (dtannx4.Rows.Count > 0)
                {
                    string LocName = objfetch.GetLocationName(Convert.ToInt32(dtWorkorder.Rows[0]["AL_LOCATION"]));
                    ExportAnnxFourData(file, Convert.ToInt16(dtWorkorder.Rows[0]["AL_LOCATION"]), Convert.ToInt16(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToString(dtWorkorder.Rows[0]["AL_TO_MONTH"]), LocName);

                }
                //  AppException.LogErrorAnnexure("Before call sp ANNEXURE 4 ", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(objAnnexure.LocationCode));
                return true;
            }


        }


        private static void ExportAnnxFourData(FileInfo ExcelCopy, int LocationCode, int YearId, string fromMonthID, string toMonthID, string LocName)
        {
            ClsFetchannexure objFetch = new ClsFetchannexure();
            string[] sOutGLCodeArray;
            string sBalType = string.Empty;
            int iRowCnt, iColCnt = 0;
            decimal bAmount = 0;
            string sFirstColGLCode;

            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            DataSet dsDetails = new DataSet();
            DataTable dtDetails = new DataTable();
            DataTable dtOutputGLCodes = new DataTable();
            DataTable dtMonthName = new DataTable();
            dtMonthName = objFetch.getmonthname(Convert.ToString(fromMonthID), Convert.ToString(toMonthID));
            decimal Total = 0;
            //  clsSession objSession;
            try
            {
                // objSession = new clsSession();
                //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                //if (objSession != null)
                //{
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 4"];
                    iRowCnt = xlSheet.Dimension.End.Row;
                    iColCnt = 46;
                    // string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);
                    //   xlSheets.Where(x => x.Name != xlSheet.ToString()).ToList().ForEach(y => xlSheets[y.Name].Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden);

                    for (int i = 11; i <= iRowCnt; i++)
                    {
                        sFirstColGLCode = xlSheet.Cells[i, 1].Value == null ? "0" : xlSheet.Cells[i, 1].Value.ToString();
                        if (sFirstColGLCode != "0")
                        {
                            if (sFirstColGLCode == "GRAND TOTAL")
                            {
                                break;
                            }
                            for (int j = 2; j <= iColCnt; j++)
                            {
                                sOutGLCodeArray = xlSheet.Cells[9, j].Value == null ? null : xlSheet.Cells[9, j].Value.ToString().Split('-').ToArray();
                                if (sOutGLCodeArray != null)
                                {
                                    //dtMonthName = objAnnexure.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);
                                    // dsDetails = objAnnexure.GetAnnexureDetails("A4", sFirstColGLCode, "0", sOutGLCodeArray[0], objAnnexure.LocationCode.ToString(), sOutGLCodeArray[1] == "C" ? "D" : "C", objAnnexure.FromDate.ToString(), objAnnexure.ToDate.ToString(), objAnnexure.FromDate.ToString(), objAnnexure.ToDate.ToString());
                                    dsDetails = objFetch.GetAnnexureDetails("A4", sFirstColGLCode, "0", sOutGLCodeArray[0], Convert.ToString(LocationCode), sOutGLCodeArray[1] == "C" ? "D" : "C", fromMonthID, toMonthID, fromMonthID, toMonthID);
                                    dtDetails = dsDetails.Tables[0];

                                    Total = 0;
                                    var InputArray = xlSheet.Cells[9, j].Value.ToString().Split(',').ToArray();

                                    // foreach (string AccCode in InputArray)
                                    // {
                                    bAmount = dtDetails.AsEnumerable().Select(val => val.Field<decimal>("Amount")).Sum();
                                    Total = bAmount;
                                    // }
                                    if (Total != 0)
                                    {
                                        xlSheet.Cells[i, j].Value = Total;
                                    }
                                }
                            }
                        }
                    }
                    using (ExcelRange rng = xlSheet.Cells["A5:D5"])
                    {
                        rng.Merge = true;
                        // xlSheet.Cells[1, 1].Value = dtMonthName.Rows[0]["YMC_Month_Name"] + " to " + dtMonthName.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.";
                        xlSheet.Cells[1, 1].Value = "MARCH 2020 FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.";
                        xlSheet.Cells[2, 1].Value = "RECONCILIATION STATEMENT OF ASSETS RELEASED FROM SERVICE DURING THE PERIOD " + dtMonthName.Rows[0]["YMC_Month_Name"] + " to " + dtMonthName.Rows[1]["YMC_Month_Name"] + " ";

                        xlSheet.Cells["A5:D5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + "   ";
                        xlSheet.Cells["A5:D5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A5:D5"].Style.Font.Bold = true;
                    }
                }
                xlPackage.Save();
                // }
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }

        private static string CreateNewFolder(string sLocCode, string ExcelFileName)
        {
            string sDestinationPath = string.Empty;
            string sSourcefile, sXlsxDestinationFile;
            //string sSourcePath = System.Web.HttpContext.Current.Server.MapPath("~/ExcelWorkbook");
            string sSourcePath = @"C:\ExcelWorkbook";

            // clsSession objSession;
            try
            {
                //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                sDestinationPath = @"C:\FinalAccounts\" + sLocCode;
                // sDestinationPath = System.Web.HttpContext.Current.Server.MapPath("~/FinalAccounts/" + sLocCode);
                sSourcefile = System.IO.Path.Combine(sSourcePath, ExcelFileName);
                sXlsxDestinationFile = System.IO.Path.Combine(sDestinationPath, ExcelFileName);

                //Create new folder w.r.t location code
                if (!System.IO.Directory.Exists(sDestinationPath))
                {
                    // AppException.LogError("folder created", Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");

                    Directory.CreateDirectory(sDestinationPath);
                }
                //Create copy of excel 
                if (System.IO.Directory.Exists(sSourcePath))
                {
                    // AppException.LogError("Source Path exists", Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");

                    if (System.IO.File.Exists(sXlsxDestinationFile))
                    {
                        System.IO.File.Delete(sXlsxDestinationFile);
                    }
                    System.IO.File.Copy(sSourcefile, sXlsxDestinationFile);
                }
                // }
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, sLocCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return sDestinationPath;
        }

        private static bool uploadToFtp(string headerfile, string strPath)
        {
            //   bool UploadSuccess = false;
            try
            {

                var filename1 = Path.GetFileName(headerfile);
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(strPath + filename1));                        ///check
                //ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpRequest.UseBinary = true;
                ftpRequest.Credentials = new NetworkCredential(username, password);
                //ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                //FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();

                //////testing

                ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
                using (FileStream fs = File.OpenRead(headerfile))
                {
                    byte[] buffer1 = new byte[fs.Length];
                    fs.Read(buffer1, 0, buffer1.Length);
                    fs.Close();
                    Stream requestStream = ftpRequest.GetRequestStream();
                    requestStream.Write(buffer1, 0, buffer1.Length);
                    requestStream.Flush();
                    requestStream.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);

                AppException.WritetoFile(ex.Message, "", System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
                return false;
            }
        }

        private static bool CreateDirectory(string strPath, string strappend)
        {
            try
            {
                bool checkExist = FtpDirectoryExists(strPath + strappend, username, password);



                if (checkExist == false)
                {
                    ftpRequest = (FtpWebRequest)WebRequest.Create(new Uri(strPath + strappend));
                    /* Log in to the FTP Server with the User Name and Password Provided */
                    ftpRequest.Credentials = new NetworkCredential(username, password);
                    /* When in doubt, use these options */
                    ftpRequest.UseBinary = true;
                    ftpRequest.UsePassive = true;
                    ftpRequest.KeepAlive = true;
                    ftpRequest.Proxy = new WebProxy();
                    ftpRequest.Proxy = null;
                    /* Specify the Type of FTP Request */
                    ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;
                    /* Establish Return Communication with the FTP Server */
                    ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                    /* Resource Cleanup */
                    ftpResponse.Close();
                    ftpRequest = null;

                }
                return true;
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, "", System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
                return false;
            }



        }



        private void ExportAnnxOneDataLatest(FileInfo ExcelCopy, DataSet dsDetails, DataTable dtmonths, String LocName)
        {
            // {

            string[] columnNamesBelow;
            int RowCount = 0, k = 9, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;


            DataTable dtDetails = new DataTable();
            DataTable dtDetails1 = new DataTable();
            DataTable dtDetails2 = new DataTable();
            DataTable dtDetails3 = new DataTable();
            DataTable dtDetails4 = new DataTable();
            DataTable dtDetails5 = new DataTable();
            // DataTable dtmonths = new DataTable();
            //DataSet dsDetails = new DataSet();
            // clsSession objSession;
            try
            {
                // objSession = new clsSession();
                //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                // if (objSession != null)
                // {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 1"];

                    // xlSheet.ProtectedRanges[100];
                    xlSheet.Cells["C7:CM7"].Clear();
                    xlSheet.Cells["B10:B17"].Clear();
                    xlSheet.Cells["B20:B23"].Clear();
                    xlSheet.Cells["B26:B26"].Clear();
                    xlSheet.Cells["B29:B32"].Clear();
                    xlSheet.Cells["B36:B36"].Clear();
                    //dsDetails = objAnnexure.GetAnnexureDetailOneLatest(objAnnexure.LocationCode, objSession.YearId, objAnnexure.frmonth, objAnnexure.tomonth);
                    //  dtmonths = objFetch.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);
                    //string LocName = objFetch.GetLocationName(objAnnexure.LocationCode);
                    dtDetails = dsDetails.Tables[0];
                    dtDetails1 = dsDetails.Tables[1];
                    dtDetails2 = dsDetails.Tables[2];
                    dtDetails3 = dsDetails.Tables[3];
                    dtDetails4 = dsDetails.Tables[4];
                    dtDetails5 = dsDetails.Tables[5];

                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    xlSheet.Cells[k, j + 1].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j + 1].Style.Font.Bold = true;
                                    xlSheet.Cells[k, j + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    continue;
                                }
                                xlSheet.Cells[k, j + 1].Value = dtDetails.Rows[i - 2][j - 1] is DBNull ? 0 : dtDetails.Rows[i - 2][j - 1];

                                xlSheet.Cells[k, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }
                    //data is binding from datatable 1 end


                    //data is binding from datatable 6 start
                    k = 17;

                    if (dtDetails5.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails1.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails5.Rows.Count;
                        for (int i = 2; i <= dtDetails5.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails5.Columns.Count; j++)
                            {

                                xlSheet.Cells[k, j + 1].Value = dtDetails5.Rows[i - 2][j - 1] is DBNull ? 0 : dtDetails5.Rows[i - 2][j - 1];


                                xlSheet.Cells[k, j + 1].Style.Font.Size = 12;
                            }
                            k++;

                        }
                    }
                    //data is binding from datatable 6 end


                    //data is binding from datatable 2 start
                    k = 20;
                    // xlSheet.Cells["B19:B22"].Clear();
                    if (dtDetails1.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails1.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails1.Rows.Count;
                        for (int i = 2; i <= dtDetails1.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails1.Columns.Count; j++)
                            {

                                xlSheet.Cells[k, j + 1].Value = dtDetails1.Rows[i - 2][j - 1] is DBNull ? 0 : dtDetails1.Rows[i - 2][j - 1];

                                xlSheet.Cells[k, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }
                    //data is binding from datatable 2 end


                    //data is binding from datatable 4 start
                    k = 26;

                    if (dtDetails3.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails1.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails3.Rows.Count;
                        for (int i = 2; i <= dtDetails3.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails3.Columns.Count; j++)
                            {

                                xlSheet.Cells[k, j + 1].Value = dtDetails3.Rows[i - 2][j - 1] is DBNull ? 0 : dtDetails3.Rows[i - 2][j - 1];

                                xlSheet.Cells[k, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }
                    //data is binding from datatable 5 end

                    //data is binding from datatable 4 start
                    k = 29;

                    if (dtDetails2.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails1.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails2.Rows.Count;
                        for (int i = 2; i <= dtDetails2.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails2.Columns.Count; j++)
                            {

                                xlSheet.Cells[k, j + 1].Value = dtDetails2.Rows[i - 2][j - 1] is DBNull ? 0 : dtDetails2.Rows[i - 2][j - 1];

                                xlSheet.Cells[k, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }
                    //data is binding from datatable 5 end

                    //data is binding from datatable 6 start
                    k = 36;

                    if (dtDetails4.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails1.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails4.Rows.Count;
                        for (int i = 2; i <= dtDetails4.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails4.Columns.Count; j++)
                            {

                                xlSheet.Cells[k, j + 1].Value = dtDetails4.Rows[i - 2][j - 1] is DBNull ? 0 : dtDetails4.Rows[i - 2][j - 1];

                                xlSheet.Cells[k, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }

                    //data is binding from datatable 7 start

                    using (ExcelRange rng = xlSheet.Cells["A3:H3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A3:H3"].Value = "CONSOLIDATED STATEMENT OF CAPITAL WORKS IN PROGRESS ACCOUNTS UNDER ACCOUNT GROUP 14 FOR THE PERIOD FROM  " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "  ";
                        xlSheet.Cells["A3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A3:H3"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["C5:E5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C5:E5"].Value = "Generated Time :" +(System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        xlSheet.Cells["C5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["C5:E5"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["A4:B4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:B4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + "  ";
                        xlSheet.Cells["A4:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A4:B4"].Style.Font.Bold = true;
                    }

                }
                xlPackage.Save();
                //  using (MemoryStream MyMemoryStream = new MemoryStream())
                // {
                // xlPackage.SaveAs(MyMemoryStream);
                //MyMemoryStream.WriteTo(Response.OutputStream);
                // Response.Flush();
                // }
                // }
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }



        public void ExportAnnxTwoData(FileInfo ExcelCopy, DataSet dsDetails, DataTable dtmonths, String LocName)
        {

            string[] columnNamesBelow;
            int RowCount = 0, k = 6, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            // DataTable dtmonths = new DataTable();


            //  = new DataSet();

            // clsSession objSession;
            try
            {
                // objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                // if (objSession != null)
                //{
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 2"];
                    //xlSheet.Cells.Clear();
                    //xlSheet.View.FreezePanes(9, 1);
                    //iRowCnt = xlSheet.Dimension.End.Row;
                    //   dsDetails = objAnnexure.GetAnnexureDetailsTwo(objAnnexure.LocationCode, objSession.YearId);
                    // dtmonths = objAnnexure.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);

                    // string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);

                    if (dsDetails.Tables[0].Rows.Count > 0)
                    {
                        xlSheet.Cells[k, 1].Value = "First Half";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        k = k + 1;
                        columnNamesBelow = (from dc in dsDetails.Tables[0].Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dsDetails.Tables[0].Rows.Count;
                        for (int i = 1; i <= dsDetails.Tables[0].Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dsDetails.Tables[0].Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    continue;
                                }

                                xlSheet.Cells[k, j].Value = dsDetails.Tables[0].Rows[i - 2][j - 1] is DBNull ? 0 : Convert.ToDouble(dsDetails.Tables[0].Rows[i - 2][j - 1]);
                                xlSheet.Cells[k, j].Style.Font.Size = 12;

                                if (j == dsDetails.Tables[0].Columns.Count - 3)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 2].Address + ":" + xlSheet.Cells[k, j - 1].Address + ")";
                                }
                                if (j == dsDetails.Tables[0].Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, j - 3].Address + ":" + xlSheet.Cells[k, j - 1].Address + ")";
                                }
                            }
                            k++;
                        }
                        //xlSheet.Cells[k, 1].Value = "Grand total-";
                        //xlSheet.Cells[k, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        //xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int m = 2; m <= dsDetails.Tables[0].Columns.Count; m++)
                        //{
                        //    xlSheet.Cells[k, m].Value = dsDetails.Tables[0].Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dsDetails.Tables[0].Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                        //    xlSheet.Cells[k, m].Style.Font.Bold = true;
                        //}


                    }
                    k = k + 1;
                    if (dsDetails.Tables[1].Rows.Count > 0)
                    {
                        xlSheet.Cells[k, 1].Value = "Second Half";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        k = k + 1;
                        columnNamesBelow = (from dc in dsDetails.Tables[1].Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dsDetails.Tables[0].Rows.Count;
                        for (int i = 1; i <= dsDetails.Tables[1].Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dsDetails.Tables[1].Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    continue;
                                }
                                xlSheet.Cells[k, j].Value = dsDetails.Tables[1].Rows[i - 2][j - 1] is DBNull ? 0 : Convert.ToDouble(dsDetails.Tables[1].Rows[i - 2][j - 1]);
                                xlSheet.Cells[k, j].Style.Font.Size = 12;

                                if (j == dsDetails.Tables[1].Columns.Count - 3)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 2].Address + ":" + xlSheet.Cells[k, j - 1].Address + ")";
                                }
                                if (j == dsDetails.Tables[1].Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, j - 3].Address + ":" + xlSheet.Cells[k, j - 1].Address + ")";
                                }
                            }
                            k++;
                        }
                        //xlSheet.Cells[k, 1].Value = "Grand total-";
                        //xlSheet.Cells[k, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        //xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int m = 2; m <= dsDetails.Tables[1].Columns.Count; m++)
                        //{
                        //    xlSheet.Cells[k, m].Value = dsDetails.Tables[1].Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dsDetails.Tables[1].Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                        //    xlSheet.Cells[k, m].Style.Font.Bold = true;
                        //}


                    }
                    k = k + 1;
                    if (dsDetails.Tables[2].Rows.Count > 0)
                    {
                        xlSheet.Cells[k, 1].Value = "Consolidated annexure-2";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        k = k + 1;
                        columnNamesBelow = (from dc in dsDetails.Tables[2].Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dsDetails.Tables[0].Rows.Count;
                        for (int i = 1; i <= dsDetails.Tables[2].Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dsDetails.Tables[2].Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    continue;
                                }
                                xlSheet.Cells[k, j].Value = dsDetails.Tables[2].Rows[i - 2][j - 1] is DBNull ? 0 : Convert.ToDouble(dsDetails.Tables[2].Rows[i - 2][j - 1]);
                                xlSheet.Cells[k, j].Style.Font.Size = 12;
                                if (j == dsDetails.Tables[2].Columns.Count - 3)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 2].Address + ":" + xlSheet.Cells[k, j - 1].Address + ")";
                                }
                                if (j == dsDetails.Tables[2].Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, j - 3].Address + ":" + xlSheet.Cells[k, j - 1].Address + ")";
                                }
                            }
                            k++;
                        }
                        //xlSheet.Cells[k, 1].Value = "Grand total-";
                        //xlSheet.Cells[k, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        //xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int m = 2; m <= dsDetails.Tables[2].Columns.Count; m++)
                        //{
                        //    xlSheet.Cells[k, m].Value = dsDetails.Tables[2].Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dsDetails.Tables[2].Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                        //    xlSheet.Cells[k, m].Style.Font.Bold = true;
                        //}



                    }

                    using (ExcelRange rng = xlSheet.Cells["B2:H2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B2:H2"].Value = "MARCH 2020 FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        xlSheet.Cells["B2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B2:H2"].Style.Font.Bold = true;
                        xlSheet.Cells["B2:H2"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["B3:H3"].Style.Font.Size = 12;
                    }
                    //using (ExcelRange rng = xlSheet.Cells["B5:H5"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B5:H5"].Value = "RECONCILIATION STATEMENT OF  DEPRECIATION PROVISION (12 Series) FOR THE PERIOD FROM " + objAnnexure.FromDate + " TO " + objAnnexure.ToDate;
                    //    xlSheet.Cells["B5:H5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B5:H5"].Style.Font.Bold = true;
                    //    xlSheet.Cells["B5:H5"].Style.Font.Size = 20;

                    //}
                    using (ExcelRange rng = xlSheet.Cells["B3:H3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B3:H3"].Value = "CATEGORISATION STATEMENT FOR THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " TO " + dtmonths.Rows[1]["YMC_Month_Name"] + " SHOWING THE AMOUNT TRANSFERRED FROM CWIP TO ASSET ACCOUNT UNDER EACH HEAD OF ACCOUNT";
                        xlSheet.Cells["B3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B3:H3"].Style.Font.Bold = true;
                        xlSheet.Cells["B2:H2"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["B3:H3"].Style.Font.Size = 12;
                    }
                    //using (ExcelRange rng = xlSheet.Cells["B4:C4"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B4:C4"].Value = "ANNEXURE - 2";
                    //    xlSheet.Cells["B4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B4:C4"].Style.Font.Bold = true;
                    //    xlSheet.Cells["B2:H2"].Style.Font.Name = "Bookman Old Style";
                    //    xlSheet.Cells["B3:H3"].Style.Font.Size = 12;

                    //}
                    using (ExcelRange rng = xlSheet.Cells["C4:G4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C4:G4"].Value = "NAME OF THE ACCOUNTING UNIT: " + LocName + " ";
                        xlSheet.Cells["C4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["C4:G4"].Style.Font.Bold = true;
                        xlSheet.Cells["C4:G4"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["C4:G4"].Style.Font.Size = 12;
                    }

                    for (int i = 1; i <= dsDetails.Tables[0].Columns.Count; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }

                }
                xlPackage.Save();
                //}
            }
            catch (Exception ex)
            {
                // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }

        public void ExportAnnx2AData(FileInfo ExcelCopy, DataTable dtDetails, DataTable dtmonths, String LocName)
        {

            string sBalType = string.Empty;

            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            //     = new DataTable();

            // clsSession objSession;
            try
            {
                //   objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                // if (objSession != null)
                // {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 2A"];

                    ///  dtDetails = objAnnexure.GetAnnexure2ADetails(objAnnexure.LocationCode.ToString(), objAnnexure.frmonth.ToString(), objAnnexure.tomonth.ToString());
                    //string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);
                    if (dtDetails.Rows.Count > 0)
                    {
                        // xlSheet.Cells[5, 2].Value = objAnnexure.LocationCode;
                        // xlSheet.Cells[5, 3].Value = LocName;
                        // xlSheet.Cells[7, 3].Value = dtDetails.Rows[0][1];
                        // xlSheet.Cells[8, 3].Value = dtDetails.Rows[1][1];
                        // xlSheet.Cells[10, 3].Value = dtDetails.Rows[2][1];
                        //// xlSheet.Cells[11, 3].Value = dtDetails.Rows[3][1];
                        // xlSheet.Cells[15, 3].Value = dtDetails.Rows[4][1];
                        //  xlSheet.Cells[5, 2].Value = objAnnexure.LocationCode;
                        // xlSheet.Cells[5, 3].Value = objSession.LocationName;
                        xlSheet.Cells[7, 3].Value = dtDetails.Rows[0][2];
                        xlSheet.Cells[7, 4].Value = dtDetails.Rows[0][3];
                        xlSheet.Cells[8, 3].Value = dtDetails.Rows[1][2];
                        xlSheet.Cells[8, 4].Value = dtDetails.Rows[1][3];
                        xlSheet.Cells[15, 3].Value = dtDetails.Rows[2][2];
                        xlSheet.Cells[15, 4].Value = dtDetails.Rows[2][3];

                    }

                }
                xlPackage.Save();
                // }
            }

            catch (Exception ex)
            {
                // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }



        public void ExportAnnxThreeeData(FileInfo ExcelCopy, DataTable dtDetails, DataTable dtmonths, String LocName)
        {

            string[] columnNamesBelow;
            int RowCount = 0, k = 8, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            // List<SelectListItem> lstMonthYear = new List<SelectListItem>();
            // TempData["MonthYear"] = lstMonthYear;
            //  DataTable dtDetails = new DataTable();
            //DataTable dtmonths = new DataTable();
            DataTable dtDetailslftacnt = new DataTable();
            int l = 127;
            //  clsSession objSession;
            try
            {
                //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                //if (objSession != null)
                // {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 3"];
                    //xlSheet.Cells.Clear();
                    //iRowCnt = xlSheet.Dimension.End.Row;
                    //dtDetails = objAnnexure.GetAnnexureDetailsThree(objAnnexure.LocationCode, objSession.YearId, objAnnexure.frmonth, objAnnexure.tomonth);
                    // dtmonths = objAnnexure.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);
                    //  string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);

                    //if (objAnnexure.LocationCode == 473)
                    //{

                    //    dtDetailslftacnt = objAnnexure.GetAnnexureDetailsThreeaccnt(objAnnexure.LocationCode, objSession.YearId, objAnnexure.frmonth, objAnnexure.tomonth);

                    //}

                    //dtDetails.Columns[0].ColumnName = "Account code under 10 series";


                    //dtDetails.Columns[1].ColumnName = "Assets Categorised and debited to 10 series during the year firstQuadrant ";
                    //dtDetails.Columns[2].ColumnName = "Assets Categorised and debited to 10 series during the year SecondQuadrant";


                    //dtDetails.Columns[3].ColumnName = "Total";
                    //dtDetails.Columns[4].ColumnName = "Assets received from other units and accepted by credit to 32.310";
                    //dtDetails.Columns[5].ColumnName = "Assets received from ESCOMs / KPTCL and accepted by credit to a/c code under 42";
                    //dtDetails.Columns[6].ColumnName = "Assets transferred to other units under 32.410";
                    //dtDetails.Columns[7].ColumnName = "Assets transferred to ESCOMs / KPTCL under a/c code 28";
                    //dtDetails.Columns[8].ColumnName = "16 series";
                    //// dtDetails.Columns[9].ColumnName = "16.2";
                    //dtDetails.Columns[9].ColumnName = "12 series";
                    //dtDetails.Columns[10].ColumnName = "77.711";
                    //dtDetails.Columns[11].ColumnName = "28&25(buyback)";
                    //dtDetails.Columns[12].ColumnName = "Assets sold to banks / financial institutions under sale & lease back transactions";
                    //dtDetails.Columns[13].ColumnName = "Reclassification of amount among 10 series due to mis-classification if any in previous years (Show debits as (+) and credits as (-))";
                    //dtDetails.Columns[14].ColumnName = "Withdrawal from 10 series due to other rectifications in respect of previous years errors";
                    //dtDetails.Columns[15].ColumnName = "Any other rectifications (Show debits as (+) and credits as (-))";
                    //dtDetails.Columns[16].ColumnName = "NET AMOUNT (2(Total)+3+4) - (5+6+7+8+10+11+13)±12 & 14";

                    //dtDetails.Columns[17].ColumnName = "OPENING BALANCE AS PER TB";
                    //dtDetails.Columns[18].ColumnName = "Total of Colum 15 & 16";
                    //dtDetails.Columns[19].ColumnName = "CLOSING BALANCE AS PER TB";
                    //dtDetails.Columns[20].ColumnName = "Difference(17 - 18)";

                    // xlSheet.View.FreezePanes(10, 2);
                    if (dtDetails.Rows.Count > 0)
                    {

                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                //if (i == 1)
                                //{
                                //    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                //    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                //    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //    continue;
                                //}
                                //k = 11;

                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1];
                                xlSheet.Cells[k, j].Style.Font.Size = 12;

                                if (j == dtDetails.Columns.Count - 4)
                                {
                                    //(4 + 5 + 6) - (7 + 8 + 9 + 10 + 11 + 12 + 13 + 15)±14 & 16
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 4].Address + ":" + xlSheet.Cells[k, 6].Address + ")-sum(" + xlSheet.Cells[k, 7].Address + ":" + xlSheet.Cells[k, 13].Address + "," + xlSheet.Cells[k, 15].Address + ")+sum(" + xlSheet.Cells[k, 14].Address + "+" + xlSheet.Cells[k, 16].Address + ")";
                                }
                                if (j == dtDetails.Columns.Count - 2)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 17].Address + ":" + xlSheet.Cells[k, 18].Address + ")";
                                }

                                if (j == dtDetails.Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 19].Address + "-" + xlSheet.Cells[k, 20].Address + ")";
                                }

                                //xlSheet.Cells[k, j + 1].Value = dtDetails.Rows[i - 2][j - 1];
                                //xlSheet.Cells[k, j].Style.Font.Size = 12;

                            }
                            k++;


                        }



                        //xlSheet.Cells[k, 1].Value = "TOTAL-";
                        //xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int m = 2; m <= 21; m++)
                        //{
                        //    xlSheet.Cells[k, m].Value = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                        //    xlSheet.Cells[k, m].Style.Font.Bold = true;
                        //}


                    }

                    if (dtDetailslftacnt.Rows.Count > 0)
                    {
                        columnNamesBelow = (from dc in dtDetailslftacnt.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetailslftacnt.Rows.Count;
                        for (int i = 1; i <= dtDetailslftacnt.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetailslftacnt.Columns.Count; j++)
                            {

                                xlSheet.Cells[l, j].Value = dtDetailslftacnt.Rows[i - 1][j - 1];
                                xlSheet.Cells[l, j].Style.Font.Size = 12;
                            }
                            l++;

                        }



                        xlSheet.Cells[l, 1].Value = "TOTAL-";
                        xlSheet.Cells[l, 1].Style.Font.Bold = true;
                        for (int m = 2; m <= 21; m++)
                        {
                            xlSheet.Cells[l, m].Value = dtDetailslftacnt.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetailslftacnt.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                            xlSheet.Cells[l, m].Style.Font.Bold = true;
                        }

                    }

                    using (ExcelRange rng = xlSheet.Cells["B2:H2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B2:H2"].Value = "MARCH  FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        xlSheet.Cells["B2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B2:H2"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["B3:H3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B3:H3"].Value = "RECONCILIATION STATEMENT OF FIXED ASSETS FOR THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + " ";
                        xlSheet.Cells["B3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B3:H3"].Style.Font.Bold = true;

                    }
                    //using (ExcelRange rng = xlSheet.Cells["B4:C4"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B4:C4"].Value = "ANNEXURE - 3";
                    //    xlSheet.Cells["B4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B4:C4"].Style.Font.Bold = true;


                    //}
                    using (ExcelRange rng = xlSheet.Cells["C4:G4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C4:G4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " ";
                        xlSheet.Cells["C4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["C4:G4"].Style.Font.Bold = true;

                    }

                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }


                }

                //dtDetails.Columns.Remove("releasetwo");
                xlPackage.Save();
                // }
            }
            catch (Exception ex)
            {
                // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }


        public void ExportAnnxThreeeDataUptoMf(FileInfo ExcelCopy, DataTable dtDetails, DataTable dtmonths, String LocName)
        {

            string[] columnNamesBelow;
            int RowCount = 0, k = 8, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            //   List<SelectListItem> lstMonthYear = new List<SelectListItem>();
            // TempData["MonthYear"] = lstMonthYear;
            // DataTable dtDetails = new DataTable();
            // DataTable dtmonths = new DataTable();
            DataTable dtDetailslftacnt = new DataTable();
            int l = 127;
            // clsSession objSession;
            try
            {
                //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                //if (objSession != null)
                //{
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 3 UptoMF"];
                    //xlSheet.Cells.Clear();
                    //iRowCnt = xlSheet.Dimension.End.Row;
                    ///    dtDetails = objAnnexure.GetAnnexureDetailsThreeupMF(objAnnexure.LocationCode, objSession.YearId, Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 6), Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 12));
                    //dtmonths = objAnnexure.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);
                    //string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);

                    //if (objAnnexure.LocationCode == 473)
                    //{

                    //    dtDetailslftacnt = objAnnexure.GetAnnexureDetailsThreeaccnt(objAnnexure.LocationCode, objSession.YearId, Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 6), Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 12));

                    //}

                    //dtDetails.Columns[0].ColumnName = "Account code under 10 series";


                    //dtDetails.Columns[1].ColumnName = "Assets Categorised and debited to 10 series during the year firstQuadrant ";
                    //dtDetails.Columns[2].ColumnName = "Assets Categorised and debited to 10 series during the year SecondQuadrant";


                    //dtDetails.Columns[3].ColumnName = "Total";
                    //dtDetails.Columns[4].ColumnName = "Assets received from other units and accepted by credit to 32.310";
                    //dtDetails.Columns[5].ColumnName = "Assets received from ESCOMs / KPTCL and accepted by credit to a/c code under 42";
                    //dtDetails.Columns[6].ColumnName = "Assets transferred to other units under 32.410";
                    //dtDetails.Columns[7].ColumnName = "Assets transferred to ESCOMs / KPTCL under a/c code 28";
                    //dtDetails.Columns[8].ColumnName = "16 series";
                    //// dtDetails.Columns[9].ColumnName = "16.2";
                    //dtDetails.Columns[9].ColumnName = "12 series";
                    //dtDetails.Columns[10].ColumnName = "77.711";
                    //dtDetails.Columns[11].ColumnName = "28&25(buyback)";
                    //dtDetails.Columns[12].ColumnName = "Assets sold to banks / financial institutions under sale & lease back transactions";
                    //dtDetails.Columns[13].ColumnName = "Reclassification of amount among 10 series due to mis-classification if any in previous years (Show debits as (+) and credits as (-))";
                    //dtDetails.Columns[14].ColumnName = "Withdrawal from 10 series due to other rectifications in respect of previous years errors";
                    //dtDetails.Columns[15].ColumnName = "Any other rectifications (Show debits as (+) and credits as (-))";
                    //dtDetails.Columns[16].ColumnName = "NET AMOUNT (2(Total)+3+4) - (5+6+7+8+10+11+13)±12 & 14";

                    //dtDetails.Columns[17].ColumnName = "OPENING BALANCE AS PER TB";
                    //dtDetails.Columns[18].ColumnName = "Total of Colum 15 & 16";
                    //dtDetails.Columns[19].ColumnName = "CLOSING BALANCE AS PER TB";
                    //dtDetails.Columns[20].ColumnName = "Difference(17 - 18)";

                    // xlSheet.View.FreezePanes(10, 2);
                    if (dtDetails.Rows.Count > 0)
                    {

                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                //if (i == 1)
                                //{
                                //    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                //    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                //    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //    continue;
                                //}
                                //k = 11;

                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1];
                                xlSheet.Cells[k, j].Style.Font.Size = 12;

                                if (j == dtDetails.Columns.Count - 4)
                                {
                                    //(4 + 5 + 6) - (7 + 8 + 9 + 10 + 11 + 12 + 13 + 15)±14 & 16
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 4].Address + ":" + xlSheet.Cells[k, 6].Address + ")-sum(" + xlSheet.Cells[k, 7].Address + ":" + xlSheet.Cells[k, 13].Address + "," + xlSheet.Cells[k, 15].Address + ")+sum(" + xlSheet.Cells[k, 14].Address + "+" + xlSheet.Cells[k, 16].Address + ")";
                                }
                                if (j == dtDetails.Columns.Count - 2)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 17].Address + ":" + xlSheet.Cells[k, 18].Address + ")";
                                }

                                if (j == dtDetails.Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 19].Address + "-" + xlSheet.Cells[k, 20].Address + ")";
                                }

                                //xlSheet.Cells[k, j + 1].Value = dtDetails.Rows[i - 2][j - 1];
                                //xlSheet.Cells[k, j].Style.Font.Size = 12;

                            }
                            k++;


                        }



                        //xlSheet.Cells[k, 1].Value = "TOTAL-";
                        //xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int m = 2; m <= 21; m++)
                        //{
                        //    xlSheet.Cells[k, m].Value = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                        //    xlSheet.Cells[k, m].Style.Font.Bold = true;
                        //}


                    }

                    if (dtDetailslftacnt.Rows.Count > 0)
                    {
                        columnNamesBelow = (from dc in dtDetailslftacnt.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetailslftacnt.Rows.Count;
                        for (int i = 1; i <= dtDetailslftacnt.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetailslftacnt.Columns.Count; j++)
                            {

                                xlSheet.Cells[l, j].Value = dtDetailslftacnt.Rows[i - 1][j - 1];
                                xlSheet.Cells[l, j].Style.Font.Size = 12;
                            }
                            l++;

                        }



                        xlSheet.Cells[l, 1].Value = "TOTAL-";
                        xlSheet.Cells[l, 1].Style.Font.Bold = true;
                        for (int m = 2; m <= 21; m++)
                        {
                            xlSheet.Cells[l, m].Value = dtDetailslftacnt.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetailslftacnt.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                            xlSheet.Cells[l, m].Style.Font.Bold = true;
                        }

                    }

                    using (ExcelRange rng = xlSheet.Cells["B2:H2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B2:H2"].Value = " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        xlSheet.Cells["B2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B2:H2"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["B3:H3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B3:H3"].Value = "RECONCILIATION STATEMENT OF FIXED ASSETS FOR THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + " ";
                        xlSheet.Cells["B3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B3:H3"].Style.Font.Bold = true;

                    }
                    //using (ExcelRange rng = xlSheet.Cells["B4:C4"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B4:C4"].Value = "ANNEXURE - 3";
                    //    xlSheet.Cells["B4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B4:C4"].Style.Font.Bold = true;


                    //}
                    using (ExcelRange rng = xlSheet.Cells["C4:G4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C4:G4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " ";
                        xlSheet.Cells["C4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["C4:G4"].Style.Font.Bold = true;

                    }

                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }


                }

                //dtDetails.Columns.Remove("releasetwo");
                xlPackage.Save();
                //  }
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        //newly added annexure3 upto march
        //newly added annexure3 upto sep start

        public void ExportAnnxThreeeDataUptoSep(FileInfo ExcelCopy, DataTable dtDetails, DataTable dtmonths, String LocName)
        {

            string[] columnNamesBelow;
            int RowCount = 0, k = 8, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            // List<SelectListItem> lstMonthYear = new List<SelectListItem>();
            //TempData["MonthYear"] = lstMonthYear;
            // DataTable dtDetails = new DataTable();
            // DataTable dtmonths = new DataTable();
            DataTable dtDetailslftacnt = new DataTable();
            int l = 127;
            // clsSession objSession;
            try
            {
                // objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                //  if (objSession != null)
                // {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 3 Upto Sep"];
                    //xlSheet.Cells.Clear();
                    //iRowCnt = xlSheet.Dimension.End.Row;
                    /// dtDetails = objAnnexure.GetAnnexureDetailsThreeuptoSep(objAnnexure.LocationCode, objSession.YearId, objAnnexure.frmonth, Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 5));
                    //  dtmonths = objAnnexure.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);
                    //string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);

                    //if (objAnnexure.LocationCode == 473)
                    //{

                    //    dtDetailslftacnt = objAnnexure.GetAnnexureDetailsThreeaccnt(objAnnexure.LocationCode, objSession.YearId, objAnnexure.frmonth, Convert.ToString(Convert.ToInt32(objAnnexure.frmonth) + 5));

                    //}

                    //dtDetails.Columns[0].ColumnName = "Account code under 10 series";


                    //dtDetails.Columns[1].ColumnName = "Assets Categorised and debited to 10 series during the year firstQuadrant ";
                    //dtDetails.Columns[2].ColumnName = "Assets Categorised and debited to 10 series during the year SecondQuadrant";


                    //dtDetails.Columns[3].ColumnName = "Total";
                    //dtDetails.Columns[4].ColumnName = "Assets received from other units and accepted by credit to 32.310";
                    //dtDetails.Columns[5].ColumnName = "Assets received from ESCOMs / KPTCL and accepted by credit to a/c code under 42";
                    //dtDetails.Columns[6].ColumnName = "Assets transferred to other units under 32.410";
                    //dtDetails.Columns[7].ColumnName = "Assets transferred to ESCOMs / KPTCL under a/c code 28";
                    //dtDetails.Columns[8].ColumnName = "16 series";
                    //// dtDetails.Columns[9].ColumnName = "16.2";
                    //dtDetails.Columns[9].ColumnName = "12 series";
                    //dtDetails.Columns[10].ColumnName = "77.711";
                    //dtDetails.Columns[11].ColumnName = "28&25(buyback)";
                    //dtDetails.Columns[12].ColumnName = "Assets sold to banks / financial institutions under sale & lease back transactions";
                    //dtDetails.Columns[13].ColumnName = "Reclassification of amount among 10 series due to mis-classification if any in previous years (Show debits as (+) and credits as (-))";
                    //dtDetails.Columns[14].ColumnName = "Withdrawal from 10 series due to other rectifications in respect of previous years errors";
                    //dtDetails.Columns[15].ColumnName = "Any other rectifications (Show debits as (+) and credits as (-))";
                    //dtDetails.Columns[16].ColumnName = "NET AMOUNT (2(Total)+3+4) - (5+6+7+8+10+11+13)±12 & 14";

                    //dtDetails.Columns[17].ColumnName = "OPENING BALANCE AS PER TB";
                    //dtDetails.Columns[18].ColumnName = "Total of Colum 15 & 16";
                    //dtDetails.Columns[19].ColumnName = "CLOSING BALANCE AS PER TB";
                    //dtDetails.Columns[20].ColumnName = "Difference(17 - 18)";

                    // xlSheet.View.FreezePanes(10, 2);
                    if (dtDetails.Rows.Count > 0)
                    {

                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                //if (i == 1)
                                //{
                                //    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                //    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                //    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //    continue;
                                //}
                                //k = 11;

                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1];
                                xlSheet.Cells[k, j].Style.Font.Size = 12;

                                if (j == dtDetails.Columns.Count - 4)
                                {
                                    //(4 + 5 + 6) - (7 + 8 + 9 + 10 + 11 + 12 + 13 + 15)±14 & 16
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 4].Address + ":" + xlSheet.Cells[k, 6].Address + ")-sum(" + xlSheet.Cells[k, 7].Address + ":" + xlSheet.Cells[k, 13].Address + "," + xlSheet.Cells[k, 15].Address + ")+sum(" + xlSheet.Cells[k, 14].Address + "+" + xlSheet.Cells[k, 16].Address + ")";
                                }
                                if (j == dtDetails.Columns.Count - 2)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 17].Address + ":" + xlSheet.Cells[k, 18].Address + ")";
                                }

                                if (j == dtDetails.Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=sum(" + xlSheet.Cells[k, 19].Address + "-" + xlSheet.Cells[k, 20].Address + ")";
                                }

                                //xlSheet.Cells[k, j + 1].Value = dtDetails.Rows[i - 2][j - 1];
                                //xlSheet.Cells[k, j].Style.Font.Size = 12;

                            }
                            k++;


                        }



                        //xlSheet.Cells[k, 1].Value = "TOTAL-";
                        //xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int m = 2; m <= 21; m++)
                        //{
                        //    xlSheet.Cells[k, m].Value = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                        //    xlSheet.Cells[k, m].Style.Font.Bold = true;
                        //}


                    }

                    if (dtDetailslftacnt.Rows.Count > 0)
                    {
                        columnNamesBelow = (from dc in dtDetailslftacnt.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetailslftacnt.Rows.Count;
                        for (int i = 1; i <= dtDetailslftacnt.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetailslftacnt.Columns.Count; j++)
                            {

                                xlSheet.Cells[l, j].Value = dtDetailslftacnt.Rows[i - 1][j - 1];
                                xlSheet.Cells[l, j].Style.Font.Size = 12;
                            }
                            l++;

                        }



                        xlSheet.Cells[l, 1].Value = "TOTAL-";
                        xlSheet.Cells[l, 1].Style.Font.Bold = true;
                        for (int m = 2; m <= 21; m++)
                        {
                            xlSheet.Cells[l, m].Value = dtDetailslftacnt.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetailslftacnt.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                            xlSheet.Cells[l, m].Style.Font.Bold = true;
                        }

                    }

                    using (ExcelRange rng = xlSheet.Cells["B2:H2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B2:H2"].Value = " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        xlSheet.Cells["B2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B2:H2"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["B3:H3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B3:H3"].Value = "RECONCILIATION STATEMENT OF FIXED ASSETS FOR THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + " ";
                        xlSheet.Cells["B3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B3:H3"].Style.Font.Bold = true;

                    }
                    //using (ExcelRange rng = xlSheet.Cells["B4:C4"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B4:C4"].Value = "ANNEXURE - 3";
                    //    xlSheet.Cells["B4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B4:C4"].Style.Font.Bold = true;


                    //}
                    using (ExcelRange rng = xlSheet.Cells["C4:G4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C4:G4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + "  ";
                        xlSheet.Cells["C4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["C4:G4"].Style.Font.Bold = true;

                    }

                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }


                }

                //dtDetails.Columns.Remove("releasetwo");
                xlPackage.Save();
                //}
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        public void ExportAnnxFiveData(FileInfo ExcelCopy, DataTable dtDetails, DataTable dtmonths, String LocName)
        {
            //To compare cells data and fetching start and end position for merging
            string start = "", end = "", header = "";
            string[] columnNamesBelow;
            //  here k is starting of rows headers data 
            int RowCount = 0, k = 5, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            int toDate;
            int yearId;
            // clsGeneral.GetMonthYearId(objAnnexure.ToDate, out toDate, out yearId);


            // DataTable dtDetails = new DataTable();

            // clsSession objSession;
            try
            {
                //  objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                //if (objSession != null)
                // {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 5"];
                    //  ExcelWorksheet xlSheet = xlSheets["Asset Categorisation"];

                    xlSheet.Cells.Clear();
                    //int    iRowCnt = xlSheet.Dimension.End.Row;
                    /// dtDetails = objAnnexure.GetAnnexureDetailsFive(objAnnexure.LocationCode, Convert.ToInt32(objAnnexure.frmonth), Convert.ToInt32(objAnnexure.tomonth), objSession.YearId);

                    dtDetails.Columns[0].ColumnName = "Six Digit account code under 12 series";
                    dtDetails.Columns[1].ColumnName = "Depreciation provision made during the year by crediting 12 series and debiting 77 series";
                    dtDetails.Columns[2].ColumnName = "Depreciation received from other units under 32.310";
                    dtDetails.Columns[3].ColumnName = "Depreciation received from ESCOMs / KPTCL under A/c code 28";
                    dtDetails.Columns[4].ColumnName = "Short provision made in previous years and now brought into account by crediting 12 series and debiting 83.6";
                    dtDetails.Columns[5].ColumnName = "Withdrawal of excess provision made in previous years by debiting 12 series and crediting 65.6";
                    dtDetails.Columns[6].ColumnName = "Depreciation provision withdrawn in respect of released / dismantled assets";
                    dtDetails.Columns[7].ColumnName = "Depreciation provision transferred to other units under 32.510";
                    dtDetails.Columns[8].ColumnName = "Depreciation provision transferred to ESCOMs / KPTCL under A/c code 42";
                    dtDetails.Columns[9].ColumnName = "Reclassification among 12 series during the year (Show Credits as (+) and Debits as (-))";

                    dtDetails.Columns[10].ColumnName = "Reclassification of mis-classification (Show Credits as (+) and Debits as (-))";
                    dtDetails.Columns[11].ColumnName = "Net (2+3+4+5+10+11) - (6+7+8+9";
                    dtDetails.Columns[12].ColumnName = "Opening Balance as per Trial balance";
                    dtDetails.Columns[13].ColumnName = "Total of Column L & M";
                    dtDetails.Columns[14].ColumnName = "Closing Balance as per Trial balance";
                    dtDetails.Columns[15].ColumnName = "Difference";
                    if (dtDetails.Rows.Count > 0)
                    {
                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    xlSheet.Cells[k, j].Style.Font.Name = "Bookman Old Style";
                                    xlSheet.Cells[k, j].Style.Font.Size = 12;
                                    continue;
                                }
                                if (j == 8 && i != 1)
                                {
                                    if (!(dtDetails.Rows[i - 2][j - 1] is DBNull))
                                    {
                                        double db = Convert.ToDouble(dtDetails.Rows[i - 2][j - 1]);
                                        double db1 = Math.Abs(db);
                                        xlSheet.Cells[k, j].Value = Math.Abs(db);
                                        xlSheet.Cells[k, j].Style.Font.Name = "Bookman Old Style";
                                        xlSheet.Cells[k, j].Style.Font.Size = 12;
                                        continue;
                                    }

                                }

                                // binding the value to the sheet  
                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 2][j - 1];
                                xlSheet.Cells[k, j].Style.Font.Size = 12;
                                xlSheet.Cells[k, j].Style.Font.Name = "Bookman Old Style";
                            }
                            k++;
                        }
                        xlSheet.Cells[k, 1].Value = "TOTAL-";
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells[k, 1].Style.Font.Size = 12;
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        for (int m = 2; m <= 16; m++)
                        {
                            try
                            {

                                double sumdb = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                                xlSheet.Cells[k, m].Style.Font.Bold = true;
                                //  if (!(sumdb == 0))
                                if (!(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull))
                                {
                                    //  xlSheet.Cells[k, m].Value = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                                    // xlSheet.Cells[k, m].Style.Font.Bold = true;

                                    double db1 = Math.Abs(sumdb);
                                    xlSheet.Cells[k, m].Value = Math.Abs(db1);
                                    xlSheet.Cells[k, m].Style.Font.Bold = true;
                                    xlSheet.Cells[k, m].Style.Font.Name = "Bookman Old Style";
                                    xlSheet.Cells[k, m].Style.Font.Size = 12;
                                    continue;
                                }
                                continue;

                                //  xlSheet.Cells[k, m].Value = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                                // xlSheet.Cells[k, m].Style.Font.Bold = true;
                            }
                            catch (Exception e)
                            {

                            }

                        }


                    }

                    using (ExcelRange rng = xlSheet.Cells["B1:H1"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B1:H1"].Value = "CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED";
                        xlSheet.Cells["B1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B1:H1"].Style.Font.Bold = true;
                        xlSheet.Cells["B1:H1"].Style.Font.Size = 12;
                        xlSheet.Cells["B1:H1"].Style.WrapText = true;
                        xlSheet.Cells["B1:H1"].Style.Font.Name = "Bookman Old Style";

                        //xlSheet.Cells["B3"].Value = "CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED";
                        //xlSheet.Cells["B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //xlSheet.Cells["B3"].Style.Font.Bold = true;
                        //xlSheet.Cells["B3"].Style.Font.Size = 14;
                        //xlSheet.Cells["B3"].Style.WrapText = true;

                    }
                    //using (ExcelRange rng = xlSheet.Cells["B5:H5"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B5:H5"].Value = "RECONCILIATION STATEMENT OF  DEPRECIATION PROVISION (12 Series) FOR THE PERIOD FROM " + objAnnexure.FromDate + " TO " + objAnnexure.ToDate;
                    //    xlSheet.Cells["B5:H5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B5:H5"].Style.Font.Bold = true;
                    //    xlSheet.Cells["B5:H5"].Style.Font.Size = 20;

                    //}
                    //using (ExcelRange rng = xlSheet.Cells["G6:G6"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["G6:G6"].Value = "ANNEXURE-5";
                    //    xlSheet.Cells["G6:G6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["G6:G6"].Style.Font.Bold = true;
                    //    xlSheet.Cells["G6:G6"].Style.Font.Size = 12;
                    //    xlSheet.Cells["G6:G6"].Style.WrapText = true;

                    //}
                    using (ExcelRange rng = xlSheet.Cells["C3:G3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C3:G3"].Value = "NAME OF THE ACCOUNTING UNIT:" + LocName + "";
                        xlSheet.Cells["C3:G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["C3:G3"].Style.Font.Bold = true;
                        xlSheet.Cells["C3:G3"].Style.Font.Size = 12;
                        xlSheet.Cells["C3:G3"].Style.WrapText = true;
                        xlSheet.Cells["C3:G3"].Style.Font.Name = "Bookman Old Style";

                    }
                    //using (ExcelRange rng = xlSheet.Cells["A4:G4"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A4:G4"].Value = "LOCATION CODE-" + objAnnexure.LocationCode;
                    //    xlSheet.Cells["A4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A4:G4"].Style.Font.Bold = true;
                    //    xlSheet.Cells["A4:G4"].Style.Font.Size = 12;
                    //    xlSheet.Cells["A3:D3"].Style.Font.Name = "Bookman Old Style";
                    //}

                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }

                }
                xlPackage.Save();
                //  }
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }

        public void ExportAnnxSixData(FileInfo ExcelCopy, DataSet dsDetails, DataTable dtmonths, String LocName)
        {
            //To compare cells data and fetching start and end position for merging
            string start = "", end = "", header = "";
            string[] columnNamesBelow;
            int RowCount = 0, k = 9, counter = 1, n = 1, lastTableCount = 0;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            ClsFetchannexure objFetch = new ClsFetchannexure();
            //DataSet dsDetails = new DataSet();
            DataTable dtDetails = new DataTable();

            // clsSession objSession;
            try
            {
                //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
                // if (objSession != null)
                // {
                //DataTable dtMonthName = objAnnexure.getmonthname(objAnnexure.frmonth, objAnnexure.tomonth);
                // objAnnexure.FromDate = Convert.ToString(dtMonthName.Rows[0]["YMC_Month_Name"]);
                // objAnnexure.ToDate = Convert.ToString(dtMonthName.Rows[1]["YMC_Month_Name"]);

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 6"];
                    xlSheet.Cells.Clear();
                    //iRowCnt = xlSheet.Dimension.End.Row;
                    //   dsDetails = objAnnexure.GetAnnexure6(objAnnexure.LocationCode.ToString(), objAnnexure.frmonth, objAnnexure.tomonth);
                    lastTableCount = dsDetails.Tables[dsDetails.Tables.Count - 1].Columns.Count;
                    // string LocName = clsGeneral.GetLocationName(objAnnexure.LocationCode);
                    for (int m = 0; m < dsDetails.Tables.Count; m++)
                    {


                        if (dsDetails.Tables[m].Rows.Count > 0)
                        {
                            if (n == (dsDetails.Tables.Count))
                            {
                                header = " Final Abstract of all the above Acc Codes ";
                                goto Found;
                            }

                            if (n % 2 == 0)
                            {
                                header = Convert.ToString(dsDetails.Tables[m].Rows[0][0]) + "-" + objFetch.GetAccountCodeDescription(Convert.ToString(dsDetails.Tables[m].Rows[0][0])) + " Abstract   ";
                            }

                            else
                            {
                                header = Convert.ToString(dsDetails.Tables[m].Rows[0][0]) + "-" + objFetch.GetAccountCodeDescription(Convert.ToString(dsDetails.Tables[m].Rows[0][0])) + " Detailed   ";

                            }
                            Found:
                            n++;

                            xlSheet.Cells[k - 2, 1, k - 2, 3].Merge = true;
                            xlSheet.Cells[k - 2, 1].Value = header;
                            xlSheet.Cells[k - 2, 1].Style.Font.Bold = true;
                        }
                        dtDetails = dsDetails.Tables[m];
                        if (dtDetails.Rows.Count > 0)
                        {
                            // remove the first column
                            if (dtDetails.Columns.Count != lastTableCount)
                            {
                                dtDetails.Columns.RemoveAt(0);
                            }

                            columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                            if (columnNamesBelow.Length == 6)
                            {
                                RowCount = dtDetails.Rows.Count;
                                for (int i = 1; i <= dtDetails.Rows.Count + 1; i++)
                                {
                                    for (int j = 1; j <= dtDetails.Columns.Count; j++)
                                    {
                                        if (i == 1)
                                        {
                                            xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                            xlSheet.Cells[k, j].Style.Font.Bold = true;
                                            xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            continue;
                                        }
                                        xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 2][j - 1];
                                        if (j == 3)
                                        {
                                            if (Convert.ToString(xlSheet.Cells[k, j].Value).Equals(Convert.ToString(xlSheet.Cells[k - 1, j].Value)))
                                            {
                                                if (start == "")
                                                {
                                                    start = xlSheet.Cells[k - 1, 3].Address;
                                                    end = xlSheet.Cells[k, 3].Address;
                                                }
                                                //If cell data is equals upto last column then those cells are merged here
                                                if (i == RowCount + 1)
                                                {
                                                    end = xlSheet.Cells[k, 3].Address;
                                                    using (ExcelRange rng = xlSheet.Cells[start + ":" + end])
                                                    {
                                                        rng.Merge = true;
                                                    }
                                                    start = "";
                                                    end = "";
                                                }
                                            }

                                            else
                                            {
                                                if (start != "")
                                                {
                                                    end = xlSheet.Cells[k - 1, 3].Address;
                                                    using (ExcelRange rng = xlSheet.Cells[start + ":" + end])
                                                    {
                                                        rng.Merge = true;
                                                    }
                                                    start = "";
                                                }
                                            }
                                        }
                                    }
                                    k++;
                                }
                                xlSheet.Cells[k, 3].Value = "TOTAL-";
                                xlSheet.Cells[k, 3].Style.Font.Bold = true;
                                xlSheet.Cells[k, 4].Value = dtDetails.Compute("SUM([" + columnNamesBelow[3] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[3] + "])", ""));
                                xlSheet.Cells[k, 4].Style.Font.Bold = true;
                                xlSheet.Cells[k, 5].Value = dtDetails.Compute("SUM([" + columnNamesBelow[4] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[4] + "])", ""));
                                xlSheet.Cells[k, 5].Style.Font.Bold = true;
                                xlSheet.Cells[k, 6].Value = dtDetails.Compute("SUM([" + columnNamesBelow[5] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[5] + "])", ""));
                                xlSheet.Cells[k, 6].Style.Font.Bold = true;
                                //xlSheet.Cells[k + 1, 7].Value = dtDetails.Compute("SUM([" + columnNamesBelow[6] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[6] + "])", ""));

                            }
                            else
                            {
                                //xlSheet.Cells[k - 2, 1, k - 2, 3].Merge = true;
                                //xlSheet.Cells[k - 2, 1].Value = header;
                                //xlSheet.Cells[k - 2, 1].Style.Font.Bold = true;
                                for (int i = 1; i <= dtDetails.Rows.Count + 1; i++)
                                {
                                    for (int j = 1; j <= dtDetails.Columns.Count; j++)
                                    {
                                        if (dtDetails.Columns.Count == 4)
                                        {

                                            if (i == 1)
                                            {
                                                xlSheet.Cells[k, j + 1].Value = columnNamesBelow[j - 1];
                                                xlSheet.Cells[k, j + 1].Style.Font.Bold = true;
                                                xlSheet.Cells[k, j + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                continue;
                                            }
                                            xlSheet.Cells[k, j + 1].Value = dtDetails.Rows[i - 2][j - 1];
                                        }
                                        else
                                        {
                                            if (i == 1)
                                            {
                                                xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                                xlSheet.Cells[k, j].Style.Font.Bold = true;
                                                xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                                continue;
                                            }
                                            xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 2][j - 1];
                                        }

                                    }
                                    k++;
                                }
                                if (dtDetails.Columns.Count != lastTableCount)
                                {
                                    xlSheet.Cells[k, 2].Value = "TOTAL-";
                                    xlSheet.Cells[k, 2].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 3].Value = dtDetails.Compute("SUM([" + columnNamesBelow[0] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[0] + "])", ""));
                                    xlSheet.Cells[k, 3].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 4].Value = dtDetails.Compute("SUM([" + columnNamesBelow[1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[1] + "])", ""));
                                    xlSheet.Cells[k, 4].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 5].Value = dtDetails.Compute("SUM([" + columnNamesBelow[2] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[2] + "])", ""));
                                    xlSheet.Cells[k, 5].Style.Font.Bold = true;
                                }
                                else
                                {
                                    xlSheet.Cells[k, 1].Value = "TOTAL-";
                                    xlSheet.Cells[k, 1].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 2].Value = dtDetails.Compute("SUM([" + columnNamesBelow[1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[1] + "])", ""));
                                    xlSheet.Cells[k, 2].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 3].Value = dtDetails.Compute("SUM([" + columnNamesBelow[2] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[2] + "])", ""));
                                    xlSheet.Cells[k, 3].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 4].Value = dtDetails.Compute("SUM([" + columnNamesBelow[3] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[3] + "])", ""));
                                    xlSheet.Cells[k, 4].Style.Font.Bold = true;
                                    xlSheet.Cells[k, 5].Value = dtDetails.Compute("SUM([" + columnNamesBelow[4] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[4] + "])", ""));
                                    xlSheet.Cells[k, 5].Style.Font.Bold = true;
                                }
                            }


                            k = k + 4;
                        }
                    }
                    using (ExcelRange rng = xlSheet.Cells["B1:H1"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B1:H1"].Value = "MARCH 2020 FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED";
                        xlSheet.Cells["B1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B1:H1"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["B3:H3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B3:H3"].Value = "CONSOLIDATED STATEMENTS OF ASSETS RELEASED FROM SERVICE DURING THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "";
                        xlSheet.Cells["B3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B3:H3"].Style.Font.Bold = true;
                    }
                    //using (ExcelRange rng = xlSheet.Cells["G4:G4"])
                    //{
                    // rng.Merge = true;
                    // xlSheet.Cells["G4:G4"].Value = "ANNEXURE-6";
                    // xlSheet.Cells["G4:G4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    // xlSheet.Cells["G4:G4"].Style.Font.Bold = true;
                    //}
                    using (ExcelRange rng = xlSheet.Cells["B4:F4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B4:F4"].Value = "NAME OF THE ACCOUNTING UNIT:-" + LocName;
                        xlSheet.Cells["B4:F4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B4:F4"].Style.Font.Bold = true;
                    }
                    //using (ExcelRange rng = xlSheet.Cells["G5:G5"])
                    //{
                    // rng.Merge = true;
                    // xlSheet.Cells["G5:G5"].Value = "LOCATION CODE-" + objAnnexure.LocationCode;
                    // xlSheet.Cells["G5:G5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    // xlSheet.Cells["G5:G5"].Style.Font.Bold = true;
                    //}

                    for (int i = 1; i <= 7; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }


                }
                xlPackage.Save();
                // }
            }
            catch (Exception ex)
            {
                // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        public void UploadAnnx22Series(int loccode, string[] Annxsubtype, int frommonth, int tomonth, int year, string LogID)
        {
            string sXlsxDestinationFileRenamed = string.Empty;
            try
            {
                string sXLsxFileName = @"C:\ExcelWorkbook\ANNEXURES_22_to_22d.xlsx";
                //string sXLsxFileName = Server.MapPath("~") + "\\ExcelWorkbook\\MARCH_FINAL_ANNEXURES.xlsx";
                string sXLsxFileNames = "ANNEXURES_22_to_22d.xlsx";
                //System.IO.DirectoryInfo di = new DirectoryInfo(@"C:\FinalAccounts\" + loccode);
                //foreach (FileInfo f in di.GetFiles())
                //{
                //    f.Delete();
                //}
                string sDestinationPath = CreateNewFolder(Convert.ToString(loccode), sXLsxFileNames);

                string sXlsxDestinationFile = System.IO.Path.Combine(sDestinationPath, sXLsxFileNames);

                FileInfo file = new FileInfo(sXlsxDestinationFile);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    //for (int i = 0; i < Annxsubtype.Length; i++)
                    //{


                    //    if (Annxsubtype[i] == "Annx22a")
                    //    {
                    //        ExportAnnxtwentytwoA(file, loccode, year, frommonth, tomonth);//jayaram
                    //    }
                    //    else if (Annxsubtype[i] == "Annx22b")
                    //    {
                    //        ExportAnnxtwentytwoB(file, loccode, year, frommonth, tomonth);//jayaram

                    //    }
                    //    else if (Annxsubtype[i] == "Annx22c")
                    //    {

                    //        ExportAnnx22CorD(file, "Annx 22 C", loccode, frommonth, tomonth, year, "36.310", "36.110", "36.210");
                    //    }
                    //    else
                    //    {

                    //        ExportAnnx22CorD(file, "Annx 22 D", loccode, frommonth, tomonth, year, "37.310", "37.110", "37.210");//jayaram
                    //    }
                    //}
                    ExportAnnxtwentytwo(file, "Annx 22", loccode, year, frommonth, tomonth);
                    ExportAnnxtwentytwoA(file, "Annx 22 A", loccode, year, frommonth, tomonth);//jayaram //MMS 
                    ExportAnnxtwentytwoB(file, "Annx 22 B", loccode, year, frommonth, tomonth);//jayaram  //mms
                    ExportAnnx22CorD(file, "Annx 22 C", loccode, frommonth, tomonth, year, "36.310", "36.110", "36.210");
                    ExportAnnx22CorD(file, "Annx 22 D", loccode, frommonth, tomonth, year, "37.310", "37.110", "37.210");//jayaram

                }
                string strPath = "ftp://" + host + "/";

                string strappend = Convert.ToString(loccode) + "/";

                bool blCheck = CreateDirectory(strPath, strappend);
                var sLocalDir = strLocalDir + "\\" + strappend;

                if (!Directory.Exists(sLocalDir))
                {
                    Directory.CreateDirectory(sLocalDir);
                }

                //  var Headerfile = @"" + strLocalDir + "\\" + Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]).Replace('/', '$') + "_" + "annexures_1_to_8" + ".csv";
                var Headerfile = @"" + strLocalDir + "\\" + "ANNEXURES_22_to_22d" + "_" + Convert.ToString(loccode).Replace('/', '$') + ".csv";

                if (blCheck == true)
                {
                    //using (var stream = File.Create(Headerfile))
                    //{
                    //    stream.wr(sXlsxDestinationFile);
                    //}
                    // strPath = "ftp://" + host + "/" + Convert.ToString(dsWorkOrder.Tables[0].Rows[i]["WH_LocationCode"])+"-" + sStoreName + "/";
                    strPath = "ftp://" + host + "/" + Convert.ToString(loccode) + "/";
                    strappend = Convert.ToString(LogID) + "/";
                    var filename1 = Path.GetFileName(Headerfile);
                    blCheck = CreateDirectory(strPath, strappend);
                    if (blCheck == true)
                    {
                        deleteDirectoryFTP(strPath + strappend, "cescmysore\ftp_fms", "Idea@123");
                        blCheck = CreateDirectory(strPath, strappend);

                        sXlsxDestinationFileRenamed = sXlsxDestinationFile.Replace("ANNEXURES_22_to_22d", "ANNEXURES_22_to_22d_" + Convert.ToString(loccode));
                        System.IO.File.Move(sXlsxDestinationFile, sXlsxDestinationFileRenamed);

                        string filePath = Convert.ToString(loccode) + "/" + strappend + "ANNEXURES_22_to_22d_" + Convert.ToString(loccode) + ".xlsx";

                        bool uploadStatus = uploadToFtp(sXlsxDestinationFileRenamed, strPath + strappend);

                        Console.WriteLine("FTP Upload status : " + uploadStatus);
                        Console.WriteLine("Upload Time  END  :  " + DateTime.Now);
                        System.Threading.Thread.Sleep(5000);
                        if (uploadStatus == true)
                        {
                            System.IO.File.Delete(sXlsxDestinationFileRenamed);
                            Console.WriteLine("File Deleted ");
                        }
                        string sSqlll = "UPDATE TBLANNEXURE_LOG SET AL_ENTRY_FLAG=2,AL_PATH='" + filePath + "',AL_UPDATED_DATE=getdate() WHERE  AL_ID=" + LogID + " ";
                        DBHelper.DBExecuteNoNQuery(sConString, sSqlll);
                        //}
                    }
                }
            }
            catch (Exception e)
            {
                System.IO.File.Delete(sXlsxDestinationFileRenamed);
                Console.WriteLine("File Deleted ");
                Console.WriteLine(e.Message);
                System.Threading.Thread.Sleep(5000);
            }

        }

        public static void UpdateStartTiming(int id)
        {
            string strQry = "UPDATE TBLANNEXURE_LOG set AL_JOB_START_ON = getdate() WHERE AL_ID  = " + id + " ";
            DBHelper.DBExecuteNoNQuery(sConString, strQry);

        }
        public void UploadAnnx19Series(int loccode, string[] Annxsubtype, int frommonth, int tomonth, int year, string LogID)
        {
            try
            {
                string sXLsxFileName = @"C:\ExcelWorkbook\ANNEXURES_19.xlsx";
                //string sXLsxFileName = Server.MapPath("~") + "\\ExcelWorkbook\\MARCH_FINAL_ANNEXURES.xlsx";
                string sXLsxFileNames = "ANNEXURES_19.xlsx";
                //   System.IO.DirectoryInfo di = new DirectoryInfo(@"C:\FinalAccounts\" + loccode);
                //foreach (FileInfo f in di.GetFiles())
                //{
                //    f.Delete();
                //}
                string sDestinationPath = CreateNewFolder(Convert.ToString(loccode), sXLsxFileNames);
                string sXlsxDestinationFile = System.IO.Path.Combine(sDestinationPath, sXLsxFileNames);





                FileInfo file = new FileInfo(sXlsxDestinationFile);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    for (int i = 0; i < Annxsubtype.Length; i++)
                    {


                        if (Annxsubtype[i] == "Annx 19")
                        {
                            ExportAnnx19(file, loccode, year, frommonth, tomonth);//jayaram
                            ExportAnnxNineteenAData(file, loccode, year, frommonth, tomonth);//jayaram

                            ExportAnnxNineteenBData(file, loccode, year, frommonth, tomonth);
                        }
                        //else if (Annxsubtype[i] == "Annx19a")
                        //{
                           

                        //}
                        //else if (Annxsubtype[i] == "Annx19b")
                        //{

                        //}
                        else if (Annxsubtype[i] == "Annx 30")
                        {
                            ExportAnnexure30(file, loccode, year, frommonth, tomonth);//jayaram
                            ExportAnnexure30A(file, loccode, year, frommonth, tomonth);

                        }
                        //else if (Annxsubtype[i] == "Annx30a")
                        //{

                           
                        //}
                        else if (Annxsubtype[i] == "Anne-40 A")
                        {

                            ExportAnnexure40A(file, loccode, year, frommonth, tomonth);//jayaram
                            ExportAnnexure40B(file, loccode, year, frommonth, tomonth);//jayaram
                        }
                        //else
                        //{
                        //    ExportAnnexure40B(file, loccode, year, frommonth, tomonth);//jayaram
                        //}
                    }


                }



                string strPath = "ftp://" + host + "/";

                string strappend = Convert.ToString(loccode) + "/";



                bool blCheck = CreateDirectory(strPath, strappend);
                var sLocalDir = strLocalDir + "\\" + strappend;

                if (!Directory.Exists(sLocalDir))
                {
                    Directory.CreateDirectory(sLocalDir);
                }

                //  var Headerfile = @"" + strLocalDir + "\\" + Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]).Replace('/', '$') + "_" + "annexures_1_to_8" + ".csv";
                var Headerfile = @"" + strLocalDir + "\\" + "ANNEXURES_19" + "_" + Convert.ToString(loccode).Replace('/', '$') + ".csv";

                if (blCheck == true)
                {

                    //using (var stream = File.Create(Headerfile))
                    //{
                    //    stream.wr(sXlsxDestinationFile);
                    //}
                    // strPath = "ftp://" + host + "/" + Convert.ToString(dsWorkOrder.Tables[0].Rows[i]["WH_LocationCode"])+"-" + sStoreName + "/";
                    strPath = "ftp://" + host + "/" + Convert.ToString(loccode) + "/";
                    strappend = Convert.ToString(LogID) + "/";
                    var filename1 = Path.GetFileName(Headerfile);
                    blCheck = CreateDirectory(strPath, strappend);
                    if (blCheck == true)
                    {
                        deleteDirectoryFTP(strPath + strappend, "FTP_USER", "Idea@2016");
                        blCheck = CreateDirectory(strPath, strappend);
                        string sXlsxDestinationFileRenamed = sXlsxDestinationFile.Replace("ANNEXURES_19", "ANNEXURES_19_" + Convert.ToString(loccode));
                        System.IO.File.Move(sXlsxDestinationFile, sXlsxDestinationFileRenamed);

                        string filePath = Convert.ToString(loccode) + "/" + strappend + "ANNEXURES_19_" + Convert.ToString(loccode) + ".xlsx";

                        uploadToFtp(sXlsxDestinationFileRenamed, strPath + strappend);

                        string sSqlll = "UPDATE TBLANNEXURE_LOG SET AL_ENTRY_FLAG=2,AL_PATH='" + filePath + "',AL_UPDATED_DATE=getdate() WHERE  AL_ID=" + LogID + " ";
                        DBHelper.DBExecuteNoNQuery(sConString, sSqlll);

                        System.IO.DirectoryInfo di = new DirectoryInfo(@"C:\FinalAccounts\" + loccode);
                        foreach (FileInfo f in di.GetFiles())
                        {
                            f.Delete();
                        }
                        //}
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                System.Threading.Thread.Sleep(5000);
            }

        }
        public void UploadAnnx31(int loccode, string[] Annxsubtype, int frommonth, int tomonth, int year, string LogID)
        {
            try
            {
                string sXLsxFileName = @"C:\ExcelWorkbook\ANNEXURES_22_to_22d.xlsx";
                //string sXLsxFileName = Server.MapPath("~") + "\\ExcelWorkbook\\MARCH_FINAL_ANNEXURES.xlsx";
                string sXLsxFileNames = "ANNEXURES_31.xlsx";
                System.IO.DirectoryInfo di = new DirectoryInfo(@"C:\FinalAccounts\" + loccode);
                foreach (FileInfo f in di.GetFiles())
                {
                    f.Delete();
                }
                string sDestinationPath = CreateNewFolder(Convert.ToString(loccode), sXLsxFileNames);
                string sXlsxDestinationFile = System.IO.Path.Combine(sDestinationPath, sXLsxFileNames);





                FileInfo file = new FileInfo(sXlsxDestinationFile);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    for (int i = 0; i < Annxsubtype.Length; i++)
                    {


                        if (Annxsubtype[i] == "Annx31")
                        {
                            ExportThirtyOne(file, loccode, year, frommonth, tomonth);//jayaram
                            //ExportAnnexure40A(file, loccode, year, frommonth, tomonth);//jayaram
                           // ExportAnnexure40B(file, loccode, year, frommonth, tomonth);//jayaram
                        }

                    }


                }



                string strPath = "ftp://" + host + "/";

                string strappend = Convert.ToString(loccode) + "/";



                bool blCheck = CreateDirectory(strPath, strappend);
                var sLocalDir = strLocalDir + "\\" + strappend;

                if (!Directory.Exists(sLocalDir))
                {
                    Directory.CreateDirectory(sLocalDir);
                }

                //  var Headerfile = @"" + strLocalDir + "\\" + Convert.ToString(dtWorkorder.Rows[0]["AL_LOCATION"]).Replace('/', '$') + "_" + "annexures_1_to_8" + ".csv";
                var Headerfile = @"" + strLocalDir + "\\" + "ANNEXURES_31" + "_" + Convert.ToString(loccode).Replace('/', '$') + ".csv";

                if (blCheck == true)
                {

                    //using (var stream = File.Create(Headerfile))
                    //{
                    //    stream.wr(sXlsxDestinationFile);
                    //}
                    // strPath = "ftp://" + host + "/" + Convert.ToString(dsWorkOrder.Tables[0].Rows[i]["WH_LocationCode"])+"-" + sStoreName + "/";
                    strPath = "ftp://" + host + "/" + Convert.ToString(loccode) + "/";
                    strappend = Convert.ToString(LogID) + "/";
                    var filename1 = Path.GetFileName(Headerfile);
                    blCheck = CreateDirectory(strPath, strappend);
                    if (blCheck == true)
                    {
                        deleteDirectoryFTP(strPath + strappend, "FTP_USER", "Idea@2016");
                        blCheck = CreateDirectory(strPath, strappend);

                        string sXlsxDestinationFileRenamed = sXlsxDestinationFile.Replace("ANNEXURES_31", "ANNEXURES_31_" + Convert.ToString(loccode));
                        System.IO.File.Move(sXlsxDestinationFile, sXlsxDestinationFileRenamed);

                        string filePath = Convert.ToString(loccode) + "/" + strappend + "ANNEXURES_31_" + Convert.ToString(loccode) + ".xlsx";

                        uploadToFtp(sXlsxDestinationFileRenamed, strPath + strappend);

                        string sSqlll = "UPDATE TBLANNEXURE_LOG SET AL_ENTRY_FLAG=2,AL_PATH='" + filePath + "',AL_UPDATED_DATE=getdate() WHERE  AL_ID=" + LogID + " ";
                        DBHelper.DBExecuteNoNQuery(sConString, sSqlll);
                    }
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                System.Threading.Thread.Sleep(5000);
            }
        }

        public void ExportAnnx22CorD(FileInfo ExcelCopy, string SheetName, int LocationCode, int frmonth, int tomonth, int year, string fAccCode, string sAccCode, string tAccCode)
        {

            ClsFetchannexure objFetch = new ClsFetchannexure();
            string title;
            if (SheetName == "Annx 22 C")
            {

                title = "ANNEXURE - 22C";
            }
            else
            {
                title = "ANNEXURE - 22D";
            }
            string[] columnNamesBelow;
            string start = "", end = "";
            int RowCount = 0, k = 10, sPos = 0;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();
            DataTable dtExpport = new DataTable();
            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets[SheetName];

                    //iRowCnt = xlSheet.Dimension.End.Row;

                    dsDetails = objFetch.GetAnnexureDetails22CorD(LocationCode, year, Convert.ToString(frmonth), Convert.ToString(tomonth), fAccCode, sAccCode, tAccCode);



                    DataTable dtDetails = dsDetails.Tables[0];
                    DataTable dtLegerBalance = dsDetails.Tables[1];
                    DataTable dtManualJv = dsDetails.Tables[2];

                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    //xlSheet.Cells["A1:C10"].Clear();

                    if (dtDetails.Rows.Count > 0)
                    {
                        sPos = k;
                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {

                                if (j == 1 || j == 2 || j == 4 || j == 7)
                                {
                                    xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Size = 10;

                                }
                                else
                                {

                                    xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1] is DBNull ? 0 : Convert.ToDecimal(dtDetails.Rows[i - 1][j - 1]);
                                    xlSheet.Cells[k, j].Style.Font.Size = 10;

                                }
                                if (j == dtDetails.Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=" + xlSheet.Cells[k, 3].Address + "+" + xlSheet.Cells[k, 5].Address + "-" + xlSheet.Cells[k, 6].Address + "-" + xlSheet.Cells[k, 8].Address + "+" + xlSheet.Cells[k, 9].Address;
                                }
                            }
                            k++;

                        }
                        // ExcelMerge(xlSheet, sPos, dtDetails.Rows.Count, 4, "D,E");
                        // ExcelMerge(xlSheet, sPos, dtDetails.Rows.Count, 7, "G,H");
                        //LatestExcelMerge(xlSheet, sPos, dtDetails.Rows.Count, 2, "A,B,C,F,I,J");
                        ExcelMergeFor22B(xlSheet, sPos, dtDetails.Rows.Count, 2, "A,B");
                        //k = 60; //Commented by Anil on 09.05.2024

                        xlSheet.Cells[k, 1].Value = "TOTAL-";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int i = 1; i <= columnNamesBelow.Length; i++)
                        //{
                        //    if (i != 1 && i != 2 && i != 4 && i != 7)
                        //    {
                        //        xlSheet.Cells[k, i].Value = dtDetails.Compute("SUM([" + columnNamesBelow[i - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[i - 1] + "])", ""));
                        //        xlSheet.Cells[k, i].Style.Font.Bold = true;
                        //    }
                        //}

                        //Added by Anil 
                        //below code added by Anil on 08/05/2024 for total column and cb and Ob and Diffrence.
                        // Define the column indices to sum and corresponding Excel cell positions
                        int[] columnIndices = { 2, 4, 5, 7, 8, 9 };
                        int[] excelCellPositions = { 3, 5, 6, 8, 9, 10 };

                        // Array to store the sums corresponding to each column index
                        double[] columnSums = new double[columnIndices.Length];

                        // Loop through each row in the DataTable
                        foreach (DataRow row in dtDetails.Rows)
                        {
                            // Sum the values in the desired columns
                            for (int i = 0; i < columnIndices.Length; i++)
                            {

                                int columnIndex = columnIndices[i];
                                if (row[columnIndex] != DBNull.Value)
                                {
                                    double cellValue = Convert.ToDouble(row[columnIndex]);
                                    columnSums[i] += cellValue;
                                }
                                else
                                {
                                    columnSums[i] += 0;
                                }
                            }
                        }

                        // Output the sums to the respective Excel cells
                        for (int i = 0; i < excelCellPositions.Length; i++)
                        {
                            int excelCellPosition = excelCellPositions[i];
                            double sumValue = columnSums[i];
                            xlSheet.Cells[k, excelCellPosition].Value = sumValue;
                        }
                        //Anil  code closed here on 08/05/2024 for total column and cb and Ob and Diffrence.


                        xlSheet.Cells[++k, 1].Value = "Opening Balance as per TB on " + dtmonths.Rows[0]["YMC_Month_Name"];
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells[k, dtDetails.Columns.Count].Value = dtLegerBalance.Rows[0][1];
                        xlSheet.Cells[++k, 1].Value = "Closing Balance as per TB on " + dtmonths.Rows[1]["YMC_Month_Name"];
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells[k, dtDetails.Columns.Count].Value = dtLegerBalance.Rows[1][1];
                        xlSheet.Cells[k, dtDetails.Columns.Count].Style.Font.Bold = true;
                        xlSheet.Cells[++k, 1].Value = "Difference";
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;

                        //Uncommented by Anil on 09.05.2024
                        xlSheet.Cells[k, dtDetails.Columns.Count].Value = (Convert.ToDecimal(xlSheet.Cells[k - 3, dtDetails.Columns.Count].Value) + Convert.ToDecimal(xlSheet.Cells[k - 2, dtDetails.Columns.Count].Value)) -( Convert.ToDecimal(xlSheet.Cells[k - 1, dtDetails.Columns.Count].Value)); //Uncommented by Anil on 09.05.2024

                        xlSheet.Cells[k, dtDetails.Columns.Count].Style.Font.Bold = true;
                    }
                    if (dtManualJv.Rows.Count > 0)
                    {
                        k = k + 3;
                        using (ExcelRange rng = xlSheet.Cells["B" + k + ":H" + k])
                        {
                            rng.Merge = true;
                            xlSheet.Cells["B" + k + ":H" + k].Value = "Details of Jv's operated manually for the head " + fAccCode;
                            xlSheet.Cells["B" + k + ":H" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            xlSheet.Cells["B" + k + ":H" + k].Style.Font.Bold = true;
                            xlSheet.Cells["B" + k + ":H" + k].Style.Font.Name = "Bookman Old Style";
                        }
                        k = k + 1;
                        columnNamesBelow = (from dc in dtManualJv.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        for (int i = 1; i <= dtManualJv.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtManualJv.Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                    xlSheet.Cells[k, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    continue;
                                }
                                xlSheet.Cells[k, j].Value = dtManualJv.Rows[i - 2][j - 1];
                                xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            }
                            k++;
                        }


                    }

                    using (ExcelRange rng = xlSheet.Cells["A1:M1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A1:M1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + "  FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED. ";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A1:M1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//SAGAR
                        xlSheet.Cells["A1:M1"].Style.Font.Bold = true;
                    }

                    //using (ExcelRange rng = xlSheet.Cells["B2:H2"])
                    //{
                    //    rng.Merge = true;

                    //    xlSheet.Cells["B2:H2"].Value = " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                    //    xlSheet.Cells["B2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["B2:H2"].Style.Font.Bold = true;
                    //    xlSheet.Cells["B2:H2"].Style.Font.Name = "Bookman Old Style";
                    //}

                    using (ExcelRange rng = xlSheet.Cells["B2:I2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B2:I2"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C " + fAccCode + " FOR THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To" + dtmonths.Rows[1]["YMC_Month_Name"] + "";
                        xlSheet.Cells["B2:I2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["B2:I2"].Style.Font.Bold = true;
                        xlSheet.Cells["B2:I2"].Style.Font.Name = "Bookman Old Style";

                    }
                    using (ExcelRange rng = xlSheet.Cells["A4:B4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:B4"].Value = title;
                        xlSheet.Cells["A4:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A4:B4"].Style.Font.Bold = true;
                        xlSheet.Cells["A4:B4"].Style.Font.Name = "Bookman Old Style";

                    }
                    using (ExcelRange rng = xlSheet.Cells["A5:E5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A5:E5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + " ";//SAGAR
                        xlSheet.Cells["A5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A5:E5"].Style.Font.Bold = true;
                        xlSheet.Cells["A5:E5"].Style.Font.Name = "Bookman Old Style";

                    }

                    using (ExcelRange rng = xlSheet.Cells["G5:J5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["G5:J5"].Value = "Generated Time :"+(System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        xlSheet.Cells["G5:J5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["G5:J5"].Style.Font.Bold = true;
                    }


                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Width = 20;
                        xlSheet.Column(i).Style.WrapText = true;

                    }

                }
                xlPackage.Save();

            }
            catch (Exception ex)
            {
                // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        private void ExcelMergeFor22B(ExcelWorksheet xlSheet, int sPos, int dtRowCount, int ColumnNo, string ColumnForMerge)
        {
            try
            {
                string[] MergeColumn = ColumnForMerge.Split(',');
                char col;
                dtRowCount = dtRowCount + sPos;
                sPos = sPos + 1;
                string start = "", end = "";
                for (int i = sPos; i <= dtRowCount; i++)
                {
                    if (Convert.ToString(xlSheet.Cells[i, ColumnNo].Value).Equals(Convert.ToString(xlSheet.Cells[i - 1, ColumnNo].Value)))
                    {
                        if (start == "")
                        {
                            start = xlSheet.Cells[i - 1, ColumnNo].Address;
                            end = xlSheet.Cells[i, ColumnNo].Address;
                        }
                        if (i == dtRowCount)
                        {
                            end = xlSheet.Cells[i, ColumnNo].Address;
                            col = Convert.ToChar(end.Substring(0, 1));
                            for (int m = 1; m <= MergeColumn.Length; m++)
                            {

                                using (ExcelRange rng = xlSheet.Cells[MergeColumn[m - 1] + start.Split(col)[1] + ":" + MergeColumn[m - 1] + end.Split(col)[1]])
                                {
                                    rng.Merge = true;
                                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }
                            }
                            //if (ColumnNo == 2)
                            //{
                            //    using (ExcelRange rng = xlSheet.Cells["K" + start.Split(col)[1] + ":" + "K" + end.Split(col)[1]])
                            //    {
                            //        rng.Formula = "=SUM(C" + start.Split(col)[1] + ":" + "C" + end.Split(col)[1] + ")" + "+" + "SUM(D" + start.Split(col)[1] + ":" + "D" + end.Split(col)[1] + ")" + "+" + "SUM(F" + start.Split(col)[1] + ":" + "F" + end.Split(col)[1] + ")" + "-" + "SUM(G" + start.Split(col)[1] + ":" + "G" + end.Split(col)[1] + ")" + "-" + "SUM(H" + start.Split(col)[1] + ":" + "H" + end.Split(col)[1] + ")" + "-" + "SUM(J" + start.Split(col)[1] + ":" + "J" + end.Split(col)[1] + ")";
                            //        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            //    }
                            //}
                        }
                    }
                    else
                    {

                        if (start != "")
                        {
                            end = xlSheet.Cells[i - 1, ColumnNo].Address;
                            col = Convert.ToChar(end.Substring(0, 1));
                            for (int m = 1; m <= MergeColumn.Length; m++)
                            {

                                using (ExcelRange rng = xlSheet.Cells[MergeColumn[m - 1] + start.Split(col)[1] + ":" + MergeColumn[m - 1] + end.Split(col)[1]])
                                {
                                    rng.Merge = true;
                                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }
                            }
                            //if (ColumnNo == 2)
                            //{
                            //    using (ExcelRange rng = xlSheet.Cells["K" + start.Split(col)[1] + ":" + "K" + end.Split(col)[1]])
                            //    {
                            //        rng.Formula = "=SUM(C" + start.Split(col)[1] + ":" + "C" + end.Split(col)[1] + ")" + "+" + "SUM(D" + start.Split(col)[1] + ":" + "D" + end.Split(col)[1] + ")" + "+" + "SUM(F" + start.Split(col)[1] + ":" + "F" + end.Split(col)[1] + ")" + "-" + "SUM(G" + start.Split(col)[1] + ":" + "G" + end.Split(col)[1] + ")" + "-" + "SUM(H" + start.Split(col)[1] + ":" + "H" + end.Split(col)[1] + ")" + "-" + "SUM(J" + start.Split(col)[1] + ":" + "J" + end.Split(col)[1] + ")";
                            //        rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            //    }
                            //}
                            start = "";
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
                //throw ex;

            }
        }
        private void ExportAnnxtwentytwoB(FileInfo ExcelCopy, string SheetName, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            string[] columnNamesBelow;
            int RowCount = 0, k = 10, counter = 1, sPos = 0;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            ClsFetchannexure objFetch = new ClsFetchannexure();
            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();

            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 22-B"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    //xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    DataTable dtDetailsdates = objFetch.GetDateFormMonth(dtmonths);
                    dsDetails = objFetch.GetAnnexureDetailTwentyTwoB(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));

                    DataSet FMSDetails = objFetch.GetAnnexureDetailTwentyTwoBinFMS(LocationCode, YearId, Convert.ToString(frmonth), Convert.ToString(tomonth));
                    dtDetails = mergedatatable22B(dsDetails.Tables[1], FMSDetails);
                    string LocName = Clsgenaral.GetLocationName(LocationCode);


                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        sPos = k;
                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1];
                                xlSheet.Cells[k, j].Style.Font.Size = 12;

                                if (j == dtDetails.Columns.Count)
                                {
                                    xlSheet.Cells[k, j].Formula = "=" + xlSheet.Cells[k, 3].Address + "+" + xlSheet.Cells[k, 4].Address + "+" + xlSheet.Cells[k, 6].Address + "-" + xlSheet.Cells[k, 7].Address + "-" + xlSheet.Cells[k, 8].Address + "-" + xlSheet.Cells[k, 10].Address + "+" + xlSheet.Cells[k, 11].Address;
                                }
                            }
                            k++;
                        }
                        ExcelMergeFor22B(xlSheet, sPos, dtDetails.Rows.Count, 2, "A,B");
                        //k = 60; //Commented by Anil on 09.05.2024 
                        xlSheet.Cells[k, 1].Value = "TOTAL-";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        //for (int i = 1; i <= columnNamesBelow.Length; i++)
                        //{
                        //    if (i != 1 && i != 2 && i != 5 && i != 9)
                        //    {
                        //        xlSheet.Cells[k, i].Value = dtDetails.Compute("SUM([" + columnNamesBelow[i - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[i - 1] + "])", ""));
                        //        xlSheet.Cells[k, i].Style.Font.Bold = true;
                        //    }
                        //}

                        //Added by Anil 
                        //below code added by Anil on 08/05/2024 for total column and cb and Ob and Diffrence.
                        // Define the column indices to sum and corresponding Excel cell positions
                        int[] columnIndices = { 2, 3, 5, 6, 7, 9, 10, 11 };
                        int[] excelCellPositions = { 3, 4, 6, 7, 8, 10, 11, 12 };

                        // Array to store the sums corresponding to each column index
                        double[] columnSums = new double[columnIndices.Length];

                        // Loop through each row in the DataTable
                        foreach (DataRow row in dtDetails.Rows)
                        {
                            // Sum the values in the desired columns
                            for (int i = 0; i < columnIndices.Length; i++)
                            {
                                int columnIndex = columnIndices[i];
                                double cellValue = Convert.ToDouble(row[columnIndex]);
                                columnSums[i] += cellValue;
                            }
                        }

                        // Output the sums to the respective Excel cells
                        for (int i = 0; i < excelCellPositions.Length; i++)
                        {
                            int excelCellPosition = excelCellPositions[i];
                            double sumValue = columnSums[i];
                            xlSheet.Cells[k, excelCellPosition].Value = sumValue;
                        }

                        //Anil  code closed here on 08/05/2024 for total column and cb and Ob and Diffrence.

                        xlSheet.Cells[++k, 1].Value = "Opening Balance as per TB on " + dtmonths.Rows[0]["YMC_Month_Name"];
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells[k, 1].Style.Font.Size = 12;
                        xlSheet.Cells[k, dtDetails.Columns.Count].Value = FMSDetails.Tables[2].Rows[0][1];
                        xlSheet.Cells[++k, 1].Value = "Closing Balance as per TB on " + dtmonths.Rows[1]["YMC_Month_Name"];
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells[k, 1].Style.Font.Size = 12;
                        xlSheet.Cells[k, dtDetails.Columns.Count].Value = FMSDetails.Tables[2].Rows[1][1];
                        xlSheet.Cells[k, dtDetails.Columns.Count].Style.Font.Bold = true;

                        //Added by Anil on 09.05.2024
                        xlSheet.Cells[++k, 1].Value = "difference";
                        xlSheet.Cells[k, 1].Style.Font.Name = "bookman old style";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        xlSheet.Cells[k, 1].Style.Font.Size = 10;
                        xlSheet.Cells[k, dtDetails.Columns.Count].Value = Convert.ToDecimal(xlSheet.Cells[k - 3, dtDetails.Columns.Count].Value) + Convert.ToDecimal(xlSheet.Cells[k - 2, dtDetails.Columns.Count].Value) - Convert.ToDecimal(xlSheet.Cells[k - 1, dtDetails.Columns.Count].Value);
                        xlSheet.Cells[k, dtDetails.Columns.Count].Style.Font.Bold = true;
                        //Added above code by Anil on 09.05.2024

                        if (FMSDetails.Tables[3].Rows.Count > 0)
                        {
                            DataTable dtManualJv = FMSDetails.Tables[3];
                            k = k + 3;
                            using (ExcelRange rng = xlSheet.Cells["B" + k + ":H" + k])
                            {
                                rng.Merge = true;
                                xlSheet.Cells["B" + k + ":H" + k].Value = "Details of Jv's operated manually for the head - 32.310";
                                xlSheet.Cells["B" + k + ":H" + k].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                xlSheet.Cells["B" + k + ":H" + k].Style.Font.Bold = true;
                                xlSheet.Cells["B" + k + ":H" + k].Style.Font.Name = "Bookman Old Style";
                            }
                            k = k + 1;
                            columnNamesBelow = (from dc in dtManualJv.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                            for (int i = 1; i <= dtManualJv.Rows.Count + 1; i++)
                            {
                                for (int j = 1; j <= dtManualJv.Columns.Count; j++)
                                {
                                    if (i == 1)
                                    {
                                        xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                        xlSheet.Cells[k, j].Style.Font.Bold = true;
                                        xlSheet.Cells[k, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        xlSheet.Cells[k, j].Style.Font.Name = "Bookman Old Style";
                                        xlSheet.Cells[k, j].Style.Font.Size = 12;
                                        continue;
                                    }
                                    xlSheet.Cells[k, j].Value = dtManualJv.Rows[i - 2][j - 1];
                                    xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                }
                                k++;
                            }


                        }
                    }

                    //using (ExcelRange rng = xlSheet.Cells["A3:E3"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A3:E3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 32.310 FROM " + dtmonths.Rows[1]["YMC_Month_Name"] + "  ";
                    //    xlSheet.Cells["A3:E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A3:E3"].Style.Font.Bold = true;
                    //}

                    //using (ExcelRange rng = xlSheet.Cells["A1:E1"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A1:E1"].Value = "FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.";
                    //    xlSheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A1:E1"].Style.Font.Bold = true;
                    //}







                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Width = 30;
                        xlSheet.Column(i).Style.WrapText = true;

                    }
                    xlSheet.Column(1).Width = 60;
                    

                    using (ExcelRange rng = xlSheet.Cells["A1:M1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A1:M1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + "  FINAL ACCOUNTS OF   CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED. ";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A1:M1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A1:M1"].Style.Font.Bold = true;
                    }


                    using (ExcelRange rng = xlSheet.Cells["A2:K2"])
                    {
                        rng.Merge = true;
                        //xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A2:K2"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + " ";
                        xlSheet.Cells["A2:K2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A2:K2"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["A5:D5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A5:D5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";//SAGAR

                        xlSheet.Cells["A5:D5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A5:D5"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["E5:G5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["E5:G5"].Value = "Generated Time :" + (System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        xlSheet.Cells["E5:G5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["E5:G5"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["A4:B4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:B4"].Value = "ANNEXURE - 22B";
                        xlSheet.Cells["A4:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A4:B4"].Style.Font.Bold = true;
                        xlSheet.Cells["A4:B4"].Style.Font.Name = "Bookman Old Style";

                    }

                    //using (ExcelRange rng = xlSheet.Cells["A3:E3"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A3:E3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 32.310 FROM  " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "  ";
                    //    xlSheet.Cells["A3:E3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A3:E3"].Style.Font.Bold = true;
                    //    xlSheet.Cells["A3:E3"].Style.Font.Size = 12;
                    //}
                    //using (ExcelRange rng = xlSheet.Cells["A5:C5"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A5:C5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                    //    xlSheet.Cells["A5:C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A5:C5"].Style.Font.Bold = true;
                    //    xlSheet.Cells["A5:C5"].Style.Font.Size = 12;
                    //}
                }
                xlPackage.Save();

            }

            catch (Exception ex)
            {
                // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }



        public DataTable mergedatatable22B(DataTable table1, DataSet table)
        {
            DataTable table2 = table.Tables[0];
            DataTable location = table.Tables[1];
            DataTable final = new DataTable();
            final.Columns.Add("LocName", typeof(string));
            final.Columns.Add("LocCode", typeof(int));
            final.Columns.Add("Amt1", typeof(double));
            final.Columns.Add("Amt2", typeof(double));
            final.Columns.Add("AccCode1", typeof(string));
            final.Columns.Add("Amt3", typeof(double));
            final.Columns.Add("Amt4", typeof(double));
            final.Columns.Add("Amt5", typeof(double));
            final.Columns.Add("AccCode2", typeof(string));
            final.Columns.Add("Amt6", typeof(double));
            final.Columns.Add("Amt7", typeof(double));
            final.Columns.Add("Total", typeof(double));
            DataTable MergedDatatable = new DataTable();

            if (table1.Rows.Count > 0 && table2.Rows.Count > 0)
            {
                table1.Merge(table2, false, MissingSchemaAction.Ignore);
                //table1.Merge(table2);

                MergedDatatable = table1;

            }
            else if (table1.Rows.Count > 0)
            {
                MergedDatatable = table1;
            }
            else
            {
                MergedDatatable = table2;
            }
            foreach (DataRow dr in MergedDatatable.Rows)
            {
                DataRow d = final.NewRow();
                d["LocName"] = dr["LocName"];
                d["LocCode"] = dr["LocCode"];
                d["Amt1"] = dr["Amt1"] is DBNull ? 0 : dr["Amt1"];
                d["Amt2"] = dr["Amt2"] is DBNull ? 0 : dr["Amt2"];
                d["AccCode1"] = dr["AccCode1"] is DBNull ? "" : dr["AccCode1"];
                d["Amt3"] = dr["Amt3"] is DBNull ? 0 : dr["Amt3"];
                d["Amt4"] = dr["Amt4"] is DBNull ? 0 : dr["Amt4"];
                d["Amt5"] = dr["Amt5"] is DBNull ? 0 : dr["Amt5"];
                d["AccCode2"] = dr["AccCode2"] is DBNull ? "" : dr["AccCode2"];
                d["Amt6"] = dr["Amt6"] is DBNull ? 0 : dr["Amt6"];
                d["Amt7"] = dr["Amt7"] is DBNull ? 0 : dr["Amt7"];
                final.Rows.Add(d);
                //DataRow d = final.NewRow();
                //d["LocName"] = dr["smname"];
                //d["LocCode"] = dr["smcode"] is DBNull? 0 : dr["smcode"];
                //d["Amt1"] = dr["amount1"] is DBNull ? 0 : dr["amount1"];
                //d["Amt2"] = dr["amount2"] is DBNull ? 0 : dr["amount2"];
                //d["AccCode1"] = dr["acccode1"] is DBNull ? "" : dr["acccode1"];
                //d["Amt3"] = dr["amount3"] is DBNull ? 0 : dr["amount3"];
                //d["Amt4"] = dr["amount4"] is DBNull ? 0 : dr["amount4"];
                //d["Amt5"] = dr["amount5"] is DBNull ? 0 : dr["amount5"];
                //d["AccCode2"] = dr["acccode2"] is DBNull ? "" : dr["acccode2"];
                //d["Amt6"] = dr["amount6"] is DBNull ? 0 : dr["amount6"];
                //d["Amt7"] = dr["amount7"] is DBNull ? 0 : dr["amount7"];
                //final.Rows.Add(d);
               
            }
            var data = (from row in final.AsEnumerable()
                        group row by new { column1 = row.Field<int>("LocCode"), column2 = row.Field<string>("AccCode1"), column3 = row.Field<string>("AccCode2") } into grp
                        orderby grp.Key.column1
                        select new
                        {
                            LocCode = grp.Key.column1,
                            AccCode1 = grp.Key.column2,
                            AccCode2 = grp.Key.column3,
                            Amt1 = grp.Sum(r => r.Field<double>("Amt1")),
                            Amt2 = grp.Sum(r => r.Field<double>("Amt2")),
                            Amt3 = grp.Sum(r => r.Field<double>("Amt3")),
                            Amt4 = grp.Sum(r => r.Field<double>("Amt4")),
                            Amt5 = grp.Sum(r => r.Field<double>("Amt5")),
                            Amt6 = grp.Sum(r => r.Field<double>("Amt6")),
                            Amt7 = grp.Sum(r => r.Field<double>("Amt7")),


                        }).ToArray();

            final.Clear();
            foreach (var dr in data)
            {
                if (dr.Amt1 != 0 || dr.Amt2 != 0 || dr.Amt3 != 0 || dr.Amt4 != 0 || dr.Amt5 != 0 || dr.Amt6 != 0 || dr.Amt7 != 0)
                {
                    DataRow d = final.NewRow();
                    var LocName = location.Select("LocCode=" + dr.LocCode);

                    int LocationCode = Convert.ToInt32(LocName.Length.ToString());
                    if (LocationCode == 1)
                    {

                        d["LocName"] = LocName[0]["LocName"].ToString(); ;
                        d["LocCode"] = dr.LocCode;
                        d["Amt1"] = dr.Amt1;
                        d["Amt2"] = dr.Amt2;
                        d["AccCode1"] = dr.AccCode1;
                        d["Amt3"] = dr.Amt3;
                        d["Amt4"] = dr.Amt4;
                        d["Amt5"] = dr.Amt5;
                        d["AccCode2"] = dr.AccCode2;
                        d["Amt6"] = dr.Amt6;
                        d["Amt7"] = dr.Amt7;
                        d["Total"] = (Convert.ToDouble(dr.Amt1) + Convert.ToDouble(dr.Amt2) + Convert.ToDouble(dr.Amt3)) - (Convert.ToDouble(dr.Amt4) + Convert.ToDouble(dr.Amt5) + Convert.ToDouble(dr.Amt6)) + Convert.ToDouble(dr.Amt7);
                        final.Rows.Add(d);
                    }
                }

            }

            return final;
        }






        //public DataTable mergedatatable22B(DataTable table1, DataSet table)
        //{
        //    DataTable table2 = table.Tables[0];
        //    DataTable location = table.Tables[1];
        //    DataTable final = new DataTable();
        //    final.Columns.Add("LocName", typeof(string));
        //    final.Columns.Add("LocCode", typeof(int));
        //    final.Columns.Add("Amt1", typeof(double));
        //    final.Columns.Add("Amt2", typeof(double));
        //    final.Columns.Add("AccCode1", typeof(string));
        //    final.Columns.Add("Amt3", typeof(double));
        //    final.Columns.Add("Amt4", typeof(double));
        //    final.Columns.Add("Amt5", typeof(double));
        //    final.Columns.Add("AccCode2", typeof(string));
        //    final.Columns.Add("Amt6", typeof(double));
        //    final.Columns.Add("Amt7", typeof(double));
        //    final.Columns.Add("Total", typeof(double));
        //    DataTable MergedDatatable = new DataTable();

        //    if (table1.Rows.Count > 0 && table2.Rows.Count > 0)
        //    {
        //        table1.Merge(table2, false, MissingSchemaAction.Ignore);
        //        //table1.Merge(table2);

        //        MergedDatatable = table1;

        //    }
        //    else if (table1.Rows.Count > 0)
        //    {
        //        MergedDatatable = table1;
        //    }
        //    else
        //    {
        //        MergedDatatable = table2;
        //    }
        //    foreach (DataRow dr in MergedDatatable.Rows)
        //    {
        //        DataRow d = final.NewRow();
        //        d["LocName"] = dr["smname"];
        //        d["LocCode"] = dr["smcode"];
        //        d["Amt1"] = dr["amount1"] is DBNull ? 0 : dr["amount1"];
        //        d["Amt2"] = dr["amount2"] is DBNull ? 0 : dr["amount2"];
        //        d["AccCode1"] = dr["acccode1"] is DBNull ? "" : dr["acccode1"];
        //        d["Amt3"] = dr["amount3"] is DBNull ? 0 : dr["amount3"];
        //        d["Amt4"] = dr["amount4"] is DBNull ? 0 : dr["amount4"];
        //        d["Amt5"] = dr["amount5"] is DBNull ? 0 : dr["amount5"];
        //        d["AccCode2"] = dr["acccode2"] is DBNull ? "" : dr["acccode2"];
        //        d["Amt6"] = dr["amount6"] is DBNull ? 0 : dr["amount6"];
        //        d["Amt7"] = dr["amount7"] is DBNull ? 0 : dr["amount7"];
        //        final.Rows.Add(d);
        //    }
        //    var data = (from row in final.AsEnumerable()
        //                group row by new { column1 = row.Field<int>("LocCode"), column2 = row.Field<string>("AccCode1"), column3 = row.Field<string>("AccCode2") } into grp
        //                orderby grp.Key.column1
        //                select new
        //                {
        //                    LocCode = grp.Key.column1,
        //                    AccCode1 = grp.Key.column2,
        //                    AccCode2 = grp.Key.column3,
        //                    Amt1 = grp.Sum(r => r.Field<double>("Amt1")),
        //                    Amt2 = grp.Sum(r => r.Field<double>("Amt2")),
        //                    Amt3 = grp.Sum(r => r.Field<double>("Amt3")),
        //                    Amt4 = grp.Sum(r => r.Field<double>("Amt4")),
        //                    Amt5 = grp.Sum(r => r.Field<double>("Amt5")),
        //                    Amt6 = grp.Sum(r => r.Field<double>("Amt6")),
        //                    Amt7 = grp.Sum(r => r.Field<double>("Amt7")),


        //                }).ToArray();

        //    final.Clear();
        //    foreach (var dr in data)
        //    {
        //        if (dr.Amt1 != 0 || dr.Amt2 != 0 || dr.Amt3 != 0 || dr.Amt4 != 0 || dr.Amt5 != 0 || dr.Amt6 != 0 || dr.Amt7 != 0)
        //        {
        //            DataRow d = final.NewRow();
        //            var LocName = location.Select("LocCode=" + dr.LocCode);
        //            d["LocName"] = LocName[0]["LocName"].ToString(); ;
        //            d["LocCode"] = dr.LocCode;
        //            d["Amt1"] = dr.Amt1;
        //            d["Amt2"] = dr.Amt2;
        //            d["AccCode1"] = dr.AccCode1;
        //            d["Amt3"] = dr.Amt3;
        //            d["Amt4"] = dr.Amt4;
        //            d["Amt5"] = dr.Amt5;
        //            d["AccCode2"] = dr.AccCode2;
        //            d["Amt6"] = dr.Amt6;
        //            d["Amt7"] = dr.Amt7;
        //            d["Total"] = (Convert.ToDouble(dr.Amt1) + Convert.ToDouble(dr.Amt2) + Convert.ToDouble(dr.Amt3)) - (Convert.ToDouble(dr.Amt4) + Convert.ToDouble(dr.Amt5) + Convert.ToDouble(dr.Amt6)) + Convert.ToDouble(dr.Amt7);
        //            final.Rows.Add(d);
        //        }

        //    }

        //    return final;
        //}


        private void ExportAnnxtwentytwo(FileInfo ExcelCopy, string SheetName, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            //string title;
            //if (SheetName == "Annx 22 A")
            //{

            //    title = "ANNEXURE - 22A";
            //}

            string[] columnNamesBelow;
            int RowCount = 0, k = 9, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            string sfromDate = string.Empty;

            ClsFetchannexure onjFetch = new ClsFetchannexure();
            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();



            try
            {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 22"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    // xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));

                    //sfromDate = objClsGeneral.GetMonthName(objAnnexure.frmonth);
                    // sfromDate = "01/" + sfromDate.Replace(' ', '/');
                    //sfromDate = clsDateUtility.FormatDBDateToSys(sfromDate); // 01/05/2018
                    //sfromDate = Convert.ToDateTime(sfromDate).ToString("MM/dd/yyyy");

                    DataTable dtDetailsdates = onjFetch.GetDateFormMonth(dtmonths);
                    dsDetails = onjFetch.GetAnnexureDetailTwentyTwoA(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    dtDetails = dsDetails.Tables[0];


                    //data is binding from datatable 1 start
          

                    using (ExcelRange rng = xlSheet.Cells["A1:L1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A1:L1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + "  FINAL ACCOUNTS OF   CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED. ";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A1:L1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A1:L1"].Style.Font.Bold = true;
                    }


                    using (ExcelRange rng = xlSheet.Cells["A2:L2"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A2:L2"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "  ";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A2:L2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A2:L2"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["A4:E4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:E4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";//SAGAR

                        xlSheet.Cells["A4:E4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A4:E4"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["H4:J4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["H4:J4"].Value = "Generated Time :" + (System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        xlSheet.Cells["H4:J4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["H4:J4"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A3:B3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A3:B3"].Value = "ANNEXURE - 22";
                        xlSheet.Cells["A3:B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A3:B3"].Style.Font.Bold = true;
                    }


                    //using (ExcelRange rng = xlSheet.Cells["B45:I45"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B45:I45"].Value = "The Opening balance & Closing Balance as per this annexure should tally to OB & CB in Final TB " + "";
                    //    xlSheet.Cells["B45:I45"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //    xlSheet.Cells["B45:I45"].Style.Font.Bold = true;
                    //}

                    //using (ExcelRange rng = xlSheet.Cells["A4:B4"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A4:B4"].Value = "ANNEXURE - 22A";
                    //    xlSheet.Cells["A4:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A4:B4"].Style.Font.Bold = true;
                    //    xlSheet.Cells["A4:B4"].Style.Font.Name = "Bookman Old Style";

                    //}


                }
                xlPackage.Save();

            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        private void ExportAnnxtwentytwoA(FileInfo ExcelCopy, string SheetName, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            //string title;
            //if (SheetName == "Annx 22 A")
            //{

            //    title = "ANNEXURE - 22A";
            //}

            string[] columnNamesBelow;
            int RowCount = 0, k = 9, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            string sfromDate = string.Empty;

            ClsFetchannexure onjFetch = new ClsFetchannexure();
            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();

         

            try
            {
                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 22 A"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    // xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));

                    //sfromDate = objClsGeneral.GetMonthName(objAnnexure.frmonth);
                   // sfromDate = "01/" + sfromDate.Replace(' ', '/');
                    //sfromDate = clsDateUtility.FormatDBDateToSys(sfromDate); // 01/05/2018
                    //sfromDate = Convert.ToDateTime(sfromDate).ToString("MM/dd/yyyy");

                    DataTable dtDetailsdates = onjFetch.GetDateFormMonth(dtmonths);
                    dsDetails = onjFetch.GetAnnexureDetailTwentyTwoA(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));

                    string GetCB = onjFetch.Get31CB(LocationCode, YearId, Convert.ToString(frmonth), Convert.ToString(tomonth));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    dtDetails = dsDetails.Tables[0];


                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                //if (i == 1)
                                //{
                                // xlSheet.Cells[k, j + 1].Value = columnNamesBelow[j - 1];
                                // xlSheet.Cells[k, j + 1].Style.Font.Bold = true;
                                //xlSheet.Cells[k, j + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //continue;
                                //}dtDetails.Rows[i - 2][j - 1];
                                xlSheet.Cells[k + 1, j + 1].Value = dtDetails.Rows[i - 1][j - 1];

                                xlSheet.Cells[k + 1, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }
                    xlSheet.Cells["M36:M36"].Value = GetCB;
                    xlSheet.Cells[k, dtDetails.Columns.Count].Style.Font.Bold = true;

                    using (ExcelRange rng = xlSheet.Cells["A1:M1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A1:M1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + "  FINAL ACCOUNTS OF   CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED. " ;
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A1:M1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A1:M1"].Style.Font.Bold = true;
                    }


                    using (ExcelRange rng = xlSheet.Cells["A2:M2"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A2:M2"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "  ";
                       // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        //xlSheet.Cells["A3:K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//SAGAR
                        xlSheet.Cells["A2:M2"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["A5:D5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A5:D5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";//SAGAR

                        xlSheet.Cells["A5:D5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A5:D5"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["H5:J5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["H5:J5"].Value = "Generated Time :"+(System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        xlSheet.Cells["H5:J5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["H5:J5"].Style.Font.Bold = true;
                    }


                    using (ExcelRange rng = xlSheet.Cells["B45:I45"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B45:I45"].Value = "The Opening balance & Closing Balance as per this annexure should tally to OB & CB in Final TB " + "";
                        xlSheet.Cells["B45:I45"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["B45:I45"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A4:B4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:B4"].Value = "ANNEXURE - 22A";
                        xlSheet.Cells["A4:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A4:B4"].Style.Font.Bold = true;
                        xlSheet.Cells["A4:B4"].Style.Font.Name = "Bookman Old Style";

                    }


                }
                xlPackage.Save();

            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        public void ExportAnnx19(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            int k = 9, col = 4;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            DataTable dtmonths = new DataTable();

            ClsFetchannexure objFetch = new ClsFetchannexure();
            DataTable dtDetails = new DataTable();


            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 19"];
                    dtDetails = objFetch.ExportAnnx19(LocationCode, frmonth, tomonth, YearId);
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);

                    if (dtDetails.Rows.Count > 0)
                    {
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            if (i == 2)
                            {
                                k++;
                            }
                            else if (i == 19)
                            {
                                k = k + 2;
                            }
                            else if (i == 11)
                            {
                                k = k + 3;
                            }


                            xlSheet.Cells[k, col].Value = dtDetails.Rows[i - 1][0] is DBNull ? 0 : dtDetails.Rows[i - 1][0];
                            xlSheet.Cells[k, col].Style.Font.Size = 12;
                            k++;

                        }

                    }


                    using (ExcelRange rng = xlSheet.Cells["A1:F1"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A1:F1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A1:F1"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A5:C5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A5:C5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["A5:C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A5:C5"].Style.Font.Bold = true;

                    }
                    using (ExcelRange rng = xlSheet.Cells["D5:E5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["D5:E5"].Value = "Generated Time :" + (System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        xlSheet.Cells["D5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["D5:E5"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A2:F2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A2:F2"].Value = "STATEMENT OF MATERIAL STOCK ACCOUNT  UNDER ACCOUNT GROUP 22.2 TO 22.6 , PERIOD FROM  " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "";
                        xlSheet.Cells["A2:F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A2:F2"].Style.Font.Bold = true;

                    }

                    using (ExcelRange rng = xlSheet.Cells["A3:C3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A3:C3"].Value = "Annx 19";
                        xlSheet.Cells["A3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A3:C3"].Style.Font.Bold = true;
                        xlSheet.Cells["A3:C3"].Style.Font.Name = "Bookman Old Style";

                    }


                    for (int i = 1; i <= dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Style.WrapText = true;
                        xlSheet.Column(i).Width = 20;
                    }

                }
                xlPackage.Save();
            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        public void ExportAnnxNineteenAData(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            string sBalType = string.Empty;

            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            ClsFetchannexure objFetch = new ClsFetchannexure();
            DataTable dtDetails = new DataTable();


            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 19-A"];

                    dtDetails = objFetch.GetAnnexure19ADetails(Convert.ToString(LocationCode), Convert.ToString(tomonth));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    DataTable dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));

                    if (dtDetails.Rows.Count > 0)
                    {
                        xlSheet.Cells[5, 4].Value = LocationCode;
                        xlSheet.Cells[5, 1].Value = LocName;
                        xlSheet.Cells[8, 2].Value = dtDetails.Rows[0][0];
                        xlSheet.Cells[8, 3].Value = dtDetails.Rows[1][0];
                        xlSheet.Cells[8, 4].Value = dtDetails.Rows[2][0];

                    } 


                    using (ExcelRange rng = xlSheet.Cells["A1:E1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["A1:E1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF   CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["A1:E1"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A5:B5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A5:B5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["A5:B5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A5:B5"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["D4:E4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["D4:E4"].Value = "ANNEXURE - 19 (A)";
                        xlSheet.Cells["D4:E4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["D4:E4"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["D5:E5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["D5:E5"].Value = "LOCATION CODE:  " + LocationCode + "  ";
                        xlSheet.Cells["D5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["D5:E5"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A2:F2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A2:F2"].Value = "STATEMENT SHOWING THE DETAILS OF BALANCE AS PER PRICING LEDGER AND GENERAL LEDGER AS ON " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "";
                        xlSheet.Cells["A2:F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A2:F2"].Style.Font.Bold = true;
                    }

                    //using (ExcelRange rng = xlSheet.Cells["A1:E1"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["A1:E1"].Value = "FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.";
                    //    xlSheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    xlSheet.Cells["A1:E1"].Style.Font.Bold = true;
                    //}

                    using (ExcelRange rng = xlSheet.Cells["A7:A7"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A7:A7"].Value = "Balance as on  " + dtmonths.Rows[1]["YMC_Month_Name"] + " ";
                        xlSheet.Cells["A7:A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A7:A7"].Style.Font.Bold = true;
                    }


                }
                xlPackage.Save();
            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        public void ExportAnnxNineteenBData(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {

            string sBalType = string.Empty;
            ClsFetchannexure objFetch = new ClsFetchannexure();
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            string[] columnNamesBelow;
            DataTable dtDetails = new DataTable();
            int RowCount = 0, k = 9, counter = 1;

            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 19-B"];
                    xlSheet.Cells.Clear();
                    dtDetails = objFetch.GetAnnexureNineteenBDetails(LocationCode.ToString(), Convert.ToString(frmonth), Convert.ToString(tomonth));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    DataTable dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    if (dtDetails.Rows.Count > 0)
                    {
                        columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count + 1; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                if (i == 1)
                                {
                                    
                                    xlSheet.Cells[k, j].Value = columnNamesBelow[j - 1];
                                    xlSheet.Cells[k, j].Style.Font.Bold = true;
                                    //xlSheet.Cells[k, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    xlSheet.Cells[k, j].Style.Font.Size = 12;
                                    xlSheet.Cells[k, j].Style.Font.Name = "Bookman Old Style";
                                    continue;
                                }
                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 2][j - 1];
                                xlSheet.Cells[k, j].Style.Font.Size = 12;
                                xlSheet.Cells[k, j].Style.Font.Name = "Bookman Old Style";
                                
                            }
                            k++;

                        }



                        xlSheet.Cells[k, 1].Value = "TOTAL-";
                        xlSheet.Cells[k, 1].Style.Font.Bold = true;
                        xlSheet.Cells[k, 1].Style.Font.Size = 12;
                        xlSheet.Cells[k, 1].Style.Font.Name = "Bookman Old Style";
                        for (int m = 3; m <= dtDetails.Columns.Count; m++)
                        {
                            xlSheet.Cells[k, m].Value = dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", "") is DBNull ? 0 : Convert.ToDouble(dtDetails.Compute("SUM([" + columnNamesBelow[m - 1] + "])", ""));
                            xlSheet.Cells[k, m].Style.Font.Bold = true;
                            xlSheet.Cells[k, m].Style.Font.Size = 12;
                            xlSheet.Cells[k, m].Style.Font.Name = "Bookman Old Style";
                        }

                    }


                    using (ExcelRange rng = xlSheet.Cells["A5:E5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A5:E5"].Style.Font.Size = 12;
                        xlSheet.Cells["A5:E5"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["A5:E5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["A5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A5:E5"].Style.Font.Bold = true;
                        
                    }
                    using (ExcelRange rng = xlSheet.Cells["G5:J5"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["G5:J5"].Style.Font.Size = 12;
                        xlSheet.Cells["G5:J5"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["G5:J5"].Value = "LOCATION CODE:  " + LocationCode + "  ";
                        xlSheet.Cells["G5:J5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["G5:J5"].Style.Font.Bold = true;
                       
                    }

                    using (ExcelRange rng = xlSheet.Cells["A2:G2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A2:G2"].Style.Font.Size = 12;
                        xlSheet.Cells["A2:G2"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["A2:G2"].Value = "STATEMENT SHOWING THE BREAK UP DETAILS FOR THE BALANCE UNDER A/c  22.320 FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "   ";
                        xlSheet.Cells["A2:G2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A2:G2"].Style.Font.Bold = true;
                      
                    }

                    using (ExcelRange rng = xlSheet.Cells["A1:G1"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A1:G1"].Style.Font.Size = 12;
                        xlSheet.Cells["A1:G1"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["A1:G1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.";
                        xlSheet.Cells["A1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A1:G1"].Style.Font.Bold = true;
                       
                    }

                    using (ExcelRange rng = xlSheet.Cells["G4:J4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["G4:J4"].Style.Font.Size = 12;
                        xlSheet.Cells["G4:J4"].Style.Font.Name = "Bookman Old Style";
                        xlSheet.Cells["G4:J4"].Value = "ANNEXURE - 19 (b)";
                        xlSheet.Cells["G4:J4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["G4:J4"].Style.Font.Bold = true;
                       
                    }
             

                    xlSheet.Cells["A:Z"].AutoFitColumns(80, 120);
                    xlSheet.Cells["A:Z"].Style.Font.Size = 12;
                    xlSheet.Cells["A:Z"].Style.Font.Name = "Bookman Old Style";
                   

                }
                xlPackage.Save();
            }


            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        private void ExportAnnexure30(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            string[] columnNamesBelow;
            int RowCount = 0, k = 10, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            ClsFetchannexure objFetch = new ClsFetchannexure();

            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();

            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 30"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    // xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    DataTable dtDetailsdates = objFetch.GetDateFormMonth(dtmonths);
                    dtDetails = objFetch.GetAnnexure30(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);


                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {

                                xlSheet.Cells[k + 1, j + 1].Value = dtDetails.Rows[i - 1][j - 1];

                                xlSheet.Cells[k + 1, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }


                    //using (ExcelRange rng = xlSheet.Cells["A1:N1"])
                    //{
                    //    rng.Merge = true;
                    //    xlSheet.Cells["B1:N1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF   CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                    //    // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                    //    xlSheet.Cells["B1:N1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                    //    xlSheet.Cells["B1:N1"].Style.Font.Bold = true;
                    //}

                    using (ExcelRange rng = xlSheet.Cells["B5:F5"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["B5:F5"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["B5:F5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["B5:F5"].Style.Font.Bold = true;
                    }
                    using (ExcelRange rng = xlSheet.Cells["A2:N2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A2:N2"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF   CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        xlSheet.Cells["A2:N2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A2:N2"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A3:N3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A3:N3"].Value = "STATEMENT SHOWING THE LIST OF RELEASED ASSETS WITH DETAILS OF TRANSACTIONS  FROM" + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "   ";
                        xlSheet.Cells["A3:N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A3:N3"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["C4:D4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["C4:D4"].Value = "LOCATION CODE:  " + LocationCode + "  ";
                        xlSheet.Cells["C4:D4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["C4:D4"].Style.Font.Bold = true;
                    }


                }
                xlPackage.Save();

            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        private void ExportAnnexure30A(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            string[] columnNamesBelow;
            int RowCount = 0, k = 13, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            ClsFetchannexure objFetch = new ClsFetchannexure();
            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();

            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx 30a"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    // xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    DataTable dtDetailsdates = objFetch.GetDateFormMonth(dtmonths);
                    dtDetails = objFetch.GetAnnexure30A(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    //dtDetails = dsDetails.Tables[0];


                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {

                                xlSheet.Cells[k + 1, j + 1].Value = dtDetails.Rows[i - 1][j - 1];

                                xlSheet.Cells[k + 1, j + 1].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }

                     using (ExcelRange rng = xlSheet.Cells["B1:N1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["B1:N1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.,";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["B1:N1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["B1:N1"].Style.Font.Bold = true;
                    }

                 

                    using (ExcelRange rng = xlSheet.Cells["A3:N3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A3:N3"].Value = "STATEMENT SHOWING THE LIST OF SCRAPPED ASSETS WITH THE DETAILS OF TRANSACTIONS DURING THE PERIOD FROM " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "   ";
                        xlSheet.Cells["A3:N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A3:N3"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["J6:L6"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["J6:L6"].Value = "LOCATION CODE:  " + LocationCode + "  ";
                        xlSheet.Cells["J6:L6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["J6:L6"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A6:E6"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A6:E6"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["A6:E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A6:E6"].Style.Font.Bold = true;
                    }

                

                }
                xlPackage.Save();

            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }

        private void ExportAnnexure40A(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            string[] columnNamesBelow;
            int RowCount = 0, k = 7, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;
            ClsFetchannexure objFetch = new ClsFetchannexure();

            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();

            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Anne-40 A"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    // xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    DataTable dtDetailsdates = objFetch.GetDateFormMonth(dtmonths);

                    dtDetails = objFetch.GetAnnexure40A(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);



                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                //if (i == 1)
                                //{
                                // xlSheet.Cells[k, j + 1].Value = columnNamesBelow[j - 1];
                                // xlSheet.Cells[k, j + 1].Style.Font.Bold = true;
                                //xlSheet.Cells[k, j + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //continue;
                                //}dtDetails.Rows[i - 2][j - 1];
                                xlSheet.Cells[k + 1, j + 2].Value = dtDetails.Rows[i - 1][j - 1];

                                xlSheet.Cells[k + 1, j + 2].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }
                    using (ExcelRange rng = xlSheet.Cells["B1:G1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["B1:G1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " Chamundeshwari Electricity Supply Corporation Limited.,";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["B1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["B1:G1"].Style.Font.Bold = true;
                    }


                    using (ExcelRange rng = xlSheet.Cells["A4:D4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:D4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["A4:D4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A4:D4"].Style.Font.Bold = true;
                    }
                      using (ExcelRange rng = xlSheet.Cells["A2:G2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A2:G2"].Value = "  Statement Showing the list of Scrap Materials held at divisional Stores From " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "   ";
                        xlSheet.Cells["A2:G2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A2:G2"].Style.Font.Bold = true;
                    }


                  
                }
                xlPackage.Save();

            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        private void ExportAnnexure40B(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            string[] columnNamesBelow;
            int RowCount = 0, k = 7, counter = 1;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            ClsFetchannexure objFetch = new ClsFetchannexure();

            DataTable dtDetails = new DataTable();
            DataSet dsDetails = new DataSet();
            DataTable dtmonths = new DataTable();

            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Anne-40 B"];

                    // xlSheet.ProtectedRanges[100];
                    //xlSheet.Cells["C7:CM7"].Clear();
                    // xlSheet.Cells["B10:B15"].Clear();
                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    DataTable dtDetailsdates = objFetch.GetDateFormMonth(dtmonths);
                    dtDetails = objFetch.GetAnnexure40B(LocationCode, YearId, Convert.ToString(dtDetailsdates.Rows[0]["dates"]), Convert.ToString(dtDetailsdates.Rows[1]["dates"]));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    //dtDetails = dsDetails.Tables[0];


                    //data is binding from datatable 1 start
                    if (dtDetails.Rows.Count > 0)
                    {
                        // columnNamesBelow = (from dc in dtDetails.Columns.Cast<DataColumn>() select dc.ColumnName).ToArray();
                        RowCount = dtDetails.Rows.Count;
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                //if (i == 1)
                                //{
                                // xlSheet.Cells[k, j + 1].Value = columnNamesBelow[j - 1];
                                // xlSheet.Cells[k, j + 1].Style.Font.Bold = true;
                                //xlSheet.Cells[k, j + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //continue;
                                //}dtDetails.Rows[i - 2][j - 1];
                                xlSheet.Cells[k + 1, j + 2].Value = dtDetails.Rows[i - 1][j - 1];

                                xlSheet.Cells[k + 1, j + 2].Style.Font.Size = 12;
                            }
                            k++;
                        }
                    }

                    using (ExcelRange rng = xlSheet.Cells["B1:G1"])
                    {
                        //rng.Merge = true;
                        xlSheet.Cells["B1:G1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " Chamundeshwari Electricity Supply Corporation Limited.,";
                        // xlSheet.Cells["A3:K3"].Value = "STATEMENT OF RECONCILATION OF BALANCE UNDER A/C 31.310 FROM  " + dtmonths.Rows[1]["YMC_Month_Name"] + "  "; //SAGAR
                        xlSheet.Cells["B1:G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;//SAGAR
                        xlSheet.Cells["B1:G1"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["A4:D4"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A4:D4"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        xlSheet.Cells["A4:D4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A4:D4"].Style.Font.Bold = true;
                    }

                    using (ExcelRange rng = xlSheet.Cells["B2:K2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["B2:K2"].Value = " Statement Showing the list of Un-Operated Materials more than 3 years held at divisional Stores From " + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "   ";
                        xlSheet.Cells["B2:K2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["B2:K2"].Style.Font.Bold = true;
                    }
                }
                xlPackage.Save();

            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }
        public void ExportThirtyOne(FileInfo ExcelCopy, int LocationCode, int YearId, int frmonth, int tomonth)
        {
            int k = 7;
            ExcelPackage xlPackage = new ExcelPackage(ExcelCopy);
            ExcelWorkbook xlBook = xlPackage.Workbook;
            ExcelWorksheets xlSheets = xlBook.Worksheets;

            DataTable dtDetails = new DataTable();
            ClsFetchannexure objFetch = new ClsFetchannexure();
            DataTable dtmonths = new DataTable();


            try
            {

                if (ExcelCopy.Exists)
                {
                    ExcelWorksheet xlSheet = xlSheets["Annx31"];

                    dtmonths = Clsgenaral.getmonthname(Convert.ToString(frmonth), Convert.ToString(tomonth));
                    DataTable dtDetailsdates = objFetch.GetDateFormMonth(dtmonths);
                    dtDetails = objFetch.GetAnnexureThirtyOne(LocationCode, YearId, Convert.ToInt32(frmonth), Convert.ToInt32(tomonth));
                    string LocName = Clsgenaral.GetLocationName(LocationCode);
                    if (dtDetails.Rows.Count > 0)
                    {
                        for (int i = 1; i <= dtDetails.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dtDetails.Columns.Count; j++)
                            {
                                xlSheet.Cells[k, j].Value = dtDetails.Rows[i - 1][j - 1] is DBNull ? 0 : dtDetails.Rows[i - 1][j - 1];
                            }
                            k++;
                        }

                    }
                    using (ExcelRange rng = xlSheet.Cells["A1:M1"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A1:M1"].Value = "" + dtmonths.Rows[1]["YMC_Month_Name"] + " FINAL ACCOUNTS OF CHAMUNDESHWARI ELECTRICITY SUPPLY CORPORATION LIMITED.";
                        xlSheet.Cells["A1:M1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A1:M1"].Style.Font.Bold = true;
                        xlSheet.Cells["A1:M1"].Style.Font.Size = 12;
                        xlSheet.Cells["A1:M1"].Style.Font.Name = "Bookman Old Style"; 

                    }
                    using (ExcelRange rng = xlSheet.Cells["A2:P2"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A2:P2"].Value = "STATEMENT SHOWING THE DETAILS OF REVENUE EXPENDITURE FINAL TRIAL BALANCE AS AGAINST BUDGET GRANTS ALLOCATED FORM" + dtmonths.Rows[0]["YMC_Month_Name"] + " To " + dtmonths.Rows[1]["YMC_Month_Name"] + "   ";
                        xlSheet.Cells["A2:P2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["A2:P2"].Style.Font.Bold = true;
                        xlSheet.Cells["A2:P2"].Style.Font.Size = 12;
                        xlSheet.Cells["A2:P2"].Style.Font.Name = "Bookman Old Style";
                    }
                    using (ExcelRange rng = xlSheet.Cells["A3:D3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["A3:D3"].Value = "NAME OF THE ACCOUNTING UNIT:  " + LocName + " - " + LocationCode + "  ";
                        //xlSheet.Cells["A3:E3"].Value = LocationCode + "-" + LocName;
                        xlSheet.Cells["A3:D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        xlSheet.Cells["A3:D3"].Style.Font.Bold = true;
                        xlSheet.Cells["A3:D3"].Style.Font.Size = 12;
                        xlSheet.Cells["A3:D3"].Style.Font.Name = "Bookman Old Style";
                    }


                    using (ExcelRange rng = xlSheet.Cells["F3:K3"])
                    {
                        rng.Merge = true;
                        xlSheet.Cells["F3:K3"].Value = "Generated Time :" + (System.DateTime.Now).ToString("dd/MM/yyyy hh:mm:ss tt");
                        //xlSheet.Cells["A3:E3"].Value = LocationCode + "-" + LocName;
                        xlSheet.Cells["F3:K3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        xlSheet.Cells["F3:K3"].Style.Font.Bold = true;
                        xlSheet.Cells["F3:K3"].Style.Font.Size = 12;
                        xlSheet.Cells["A3:D3"].Style.Font.Name = "Bookman Old Style";
                    }
                    for (int i = 1; i < dtDetails.Columns.Count; i++)
                    {
                        xlSheet.Column(i).Width = 30;
                    }

                }
                xlPackage.Save();
            }

            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
        }

    }
}
