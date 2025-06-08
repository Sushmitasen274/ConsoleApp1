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

namespace MMS_FMS.Integration
{
    using System;
    using System.Threading.Tasks;

    class Program
    {
        // Instance variables
        string folderPath = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\excelworkbook\files";
        string outputFile = @"E:\ConsoleApp\FMS console Final by FMSA3\FMS console Final by FMSA3\excelworkbook\consolidated_inventory.xlsx";

        // The static Main method (entry point)
        static async Task Main(string[] args)
        {
            try
            {
                // Print start message
                Console.WriteLine("Start");

                // Create an instance of Program class to access instance variables
                var main = new Program();

                // Create the consolidator object and call the consolidation method
                var consolidator = new InventoryConsolidator();
                await consolidator.ConsolidateInventoryDataAsync(main.folderPath, main.outputFile);

                // Print completion message
                Console.WriteLine("Consolidation complete!");
            }
            catch (Exception ex)
            {
                // Handle any errors
                Console.WriteLine(ex.Message);
                Console.WriteLine("Error during consolidation.");
            }
        }
    






    //    static void Main(string[] args)
    //    {
    //        try
    //        {
    //            Console.WriteLine("STart");

    //            UploadEstimateData();
    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine(ex.Message);
    //            // System.Threading.Thread.Sleep(50000);
    //            Console.WriteLine("STart");
    //        }

    //    }
}
    //class Program
    //{
    //    private static string host = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPFILEPATH"]);
    //    private static string username = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPUSER"]);
    //    private static string password = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPPASSWORD"]);
    //    private static string port = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FTPPORT"]);
    //    private static string strLocalDir = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["LOCALDIRECTORY"]);
    //    private static FtpWebRequest ftpRequest = null;
    //    private static FtpWebResponse ftpResponse = null;

    //    static void Main(string[] args)
    //    {
    //        try
    //        {
    //            Console.WriteLine("STart");

    //            UploadEstimateData();
    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine(ex.Message);
    //            // System.Threading.Thread.Sleep(50000);
    //            Console.WriteLine("STart");
    //        }

    //    }

    //    public static void UploadEstimateData()
    //    {


    //        string strQry = string.Empty;
    //        try
    //        {
    //            // Console.WriteLine("beforehitDB");
    //            string ConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
    //            //Console.WriteLine("aftrhitDB");
    //            string strWoNO = string.Empty;
    //            string strWoDate = string.Empty;
    //            string strPath = string.Empty;
    //            string strappend = string.Empty;
    //            string sStoreName = string.Empty;
    //            bool blStatus = false;
    //            bool checkValidation = false;

    //            int querts = 0;
    //            int quertss = 0;
    //            string strLogDirectory = Convert.ToString(System.Configuration.ConfigurationSettings.AppSettings["FMS_MMSLOG"]);


    //            //strQry = "SELECT AL_ID, AL_LOCATION, AL_FROM_MONTH, AL_TO_MONTH,AL_ANNEXURE_TYPE,AL_ANNEXURE_SUBTYPE,AL_YEARID FROM TBLANNEXURE_LOG WHERE AL_ENTRY_FLAG = 1 "
    //            //+ " and AL_ANNEXURE_TYPE in ('A3','A4','A5')  ";
    //            // +" and al_id=3456 ";


    //            strQry = " SELECT top 1 AL_ID, AL_LOCATION, AL_FROM_MONTH, AL_TO_MONTH,AL_ANNEXURE_TYPE,AL_ANNEXURE_SUBTYPE,AL_YEARID FROM TBLANNEXURE_LOG WHERE AL_ENTRY_FLAG = 1 "
    //           + " and AL_ANNEXURE_TYPE='A3'  ORDER BY AL_ID ASC  ";


    //            DataTable dtWorkorder = DBHelper.DBExecDataTable(ConString, strQry);
    //            if (dtWorkorder.Rows.Count > 0)
    //            {
    //                ClsAnnexure.UpdateStartTiming(Convert.ToInt32( dtWorkorder.Rows[0]["AL_ID"]));
    //            }
    //            Console.WriteLine("Task  START");


    //            //DataTable dtWorkorder = new DataTable();

    //            //    foreach (DataRow drDetails in dsWorkOrder.Tables[0].Rows) {
    //            ClsAnnexure clsanne = new ClsAnnexure();
    //            //if (Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_TYPE"]) == "A1")
    //            //{
    //            //    clsanne.GetAnnexure(dtWorkorder);
    //            //}
    //            if (Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_TYPE"]) == "A3")
    //            {
    //                Console.WriteLine("Running UploadAnnx22Series START  ");

    //                string Annxsubtype = Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_SUBTYPE"]);
    //                string[] arr = Annxsubtype.Split(',');
    //                Console.WriteLine("Upload Time  START  :  " + DateTime.Now);

    //                clsanne.UploadAnnx22Series(Convert.ToInt32(dtWorkorder.Rows[0]["AL_LOCATION"]), arr, Convert.ToInt32(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToInt32(dtWorkorder.Rows[0]["AL_TO_MONTH"]), Convert.ToInt32(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_ID"]));
    //                Console.WriteLine("Running UploadAnnx22Series END  ");
    //                System.Threading.Thread.Sleep(5000);

    //            }
    //            //else if (Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_TYPE"]) == "A4")
    //            //{
    //            //    string Annxsubtype = Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_SUBTYPE"]);
    //            //    string[] arr = Annxsubtype.Split(',');
    //            //    clsanne.UploadAnnx19Series(Convert.ToInt32(dtWorkorder.Rows[0]["AL_LOCATION"]), arr, Convert.ToInt32(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToInt32(dtWorkorder.Rows[0]["AL_TO_MONTH"]), Convert.ToInt32(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_ID"]));
    //            //}
    //            //else if (Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_TYPE"]) == "A5")
    //            //{
    //            //    string Annxsubtype = Convert.ToString(dtWorkorder.Rows[0]["AL_ANNEXURE_SUBTYPE"]);
    //            //    string[] arr = Annxsubtype.Split(',');
    //            //    clsanne.UploadAnnx31(Convert.ToInt32(dtWorkorder.Rows[0]["AL_LOCATION"]), arr, Convert.ToInt32(dtWorkorder.Rows[0]["AL_FROM_MONTH"]), Convert.ToInt32(dtWorkorder.Rows[0]["AL_TO_MONTH"]), Convert.ToInt32(dtWorkorder.Rows[0]["AL_YEARID"]), Convert.ToString(dtWorkorder.Rows[0]["AL_ID"]));
    //            //}



    //            // }

    //            System.IO.DirectoryInfo di = new DirectoryInfo(strLocalDir);
    //            foreach (FileInfo file in di.GetFiles())
    //            {
    //                file.Delete();
    //            }
    //            foreach (DirectoryInfo dir in di.GetDirectories())
    //            {
    //                dir.Delete(true);
    //            }
    //            Console.WriteLine("Task  END  ");

    //            // System.Threading.Thread.Sleep(500);
    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine(ex.Message);
    //            Console.WriteLine(ex.StackTrace);
    //            //System.Threading.Thread.Sleep(50000);


    //            // AppException.LogError(ex.Message, "", MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, strQry, "");

    //        }
    //    }


    //    private static string CreateNewFolder(string sLocCode, string ExcelFileName)
    //    {
    //        string sDestinationPath = string.Empty;
    //        string sSourcefile, sXlsxDestinationFile;
    //        //string sSourcePath = System.Web.HttpContext.Current.Server.MapPath("~/ExcelWorkbook");
    //        string sSourcePath = @"C:\ExcelWorkbook";

    //        // clsSession objSession;
    //        try
    //        {
    //            //objSession = (CheckSessionTimeOut.IsSessionExpired(out objSession) == true ? objSession : null);
    //            sDestinationPath = @"C:\FinalAccounts\" + sLocCode;
    //            // sDestinationPath = System.Web.HttpContext.Current.Server.MapPath("~/FinalAccounts/" + sLocCode);
    //            sSourcefile = System.IO.Path.Combine(sSourcePath, ExcelFileName);
    //            sXlsxDestinationFile = System.IO.Path.Combine(sDestinationPath, ExcelFileName);

    //            //Create new folder w.r.t location code
    //            if (!System.IO.Directory.Exists(sDestinationPath))
    //            {
    //                // AppException.LogError("folder created", Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");

    //                Directory.CreateDirectory(sDestinationPath);
    //            }
    //            //Create copy of excel 
    //            if (System.IO.Directory.Exists(sSourcePath))
    //            {
    //                // AppException.LogError("Source Path exists", Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");

    //                if (System.IO.File.Exists(sXlsxDestinationFile))
    //                {
    //                    System.IO.File.Delete(sXlsxDestinationFile);
    //                }
    //                System.IO.File.Copy(sSourcefile, sXlsxDestinationFile);
    //            }
    //            // }
    //        }
    //        catch (Exception ex)
    //        {
    //            // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
    //        }
    //        return sDestinationPath;
    //    }
    //    private static void WriteLogFile(string strPath, string sFileName, string strWorkorderId, string strWorkorderType, string strDescription)
    //    {
    //        try
    //        {

    //            if (!Directory.Exists(strPath + DateTime.Now.ToString("ddMMyyyy")))
    //            {
    //                Directory.CreateDirectory(strPath + DateTime.Now.ToString("ddMMyyyy"));
    //            }
    //            string sFilePath = strPath + DateTime.Now.ToString("ddMMyyyy") + "//" + sFileName + ".csv";
    //            if (!File.Exists(sFilePath))
    //            {
    //                File.AppendAllText(sFilePath, "WO_ID,WO_TYPE,DESCRIPTION" + Environment.NewLine);
    //            }
    //            File.AppendAllText(sFilePath, strWorkorderId + ", " + strWorkorderType + " , " + strDescription + Environment.NewLine);

    //        }

    //        catch (Exception ex)
    //        {

    //        }


    //    }

    //    private static bool uploadToFtp(string headerfile, string strPath)
    //    {
    //        //   bool UploadSuccess = false;
    //        try
    //        {

    //            var filename1 = Path.GetFileName(headerfile);
    //            ftpRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(strPath + filename1));                        ///check
    //            //ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
    //            ftpRequest.UseBinary = true;
    //            ftpRequest.Credentials = new NetworkCredential(username, password);
    //            //ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
    //            //FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();

    //            //////testing

    //            ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
    //            using (FileStream fs = File.OpenRead(headerfile))
    //            {
    //                byte[] buffer1 = new byte[fs.Length];
    //                fs.Read(buffer1, 0, buffer1.Length);
    //                fs.Close();
    //                Stream requestStream = ftpRequest.GetRequestStream();
    //                requestStream.Write(buffer1, 0, buffer1.Length);
    //                requestStream.Flush();
    //                requestStream.Close();
    //            }
    //            return true;
    //        }
    //        catch (Exception ex)
    //        {
    //            Console.WriteLine(ex.StackTrace);

    //            //  AppException.LogError(ex.Message, "", MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
    //            return false;
    //        }
    //    }

    //    private static string table_to_csv(DataTable table)
    //    {
    //        string file = "";
    //        try
    //        {
    //            foreach (DataColumn col in table.Columns)

    //                file = string.Concat(file, col.ColumnName, ",");

    //            file = file.Remove(file.LastIndexOf(','), 1);
    //            file = string.Concat(file, "\r\n");


    //            foreach (DataRow row in table.Rows)
    //            {
    //                foreach (object item in row.ItemArray)
    //                    file = string.Concat(file, item.ToString(), ",");

    //                file = file.Remove(file.LastIndexOf(','), 1);
    //                file = string.Concat(file, "\r\n");
    //            }

    //            return file;
    //        }
    //        catch (Exception ex)
    //        {
    //            // AppException.LogError(ex.Message, "", MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
    //            return file;
    //        }
    //    }

    //    private static bool CreateDirectory(string strPath, string strappend)
    //    {
    //        try
    //        {
    //            bool checkExist = FtpDirectoryExists(strPath + strappend, username, password);



    //            if (checkExist == false)
    //            {
    //                ftpRequest = (FtpWebRequest)WebRequest.Create(new Uri(strPath + strappend));
    //                /* Log in to the FTP Server with the User Name and Password Provided */
    //                ftpRequest.Credentials = new NetworkCredential(username, password);
    //                /* When in doubt, use these options */
    //                ftpRequest.UseBinary = true;
    //                ftpRequest.UsePassive = true;
    //                ftpRequest.KeepAlive = true;
    //                ftpRequest.Proxy = new WebProxy();
    //                ftpRequest.Proxy = null;
    //                /* Specify the Type of FTP Request */
    //                ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;
    //                /* Establish Return Communication with the FTP Server */
    //                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
    //                /* Resource Cleanup */
    //                ftpResponse.Close();
    //                ftpRequest = null;

    //            }
    //            return true;
    //        }
    //        catch (Exception ex)
    //        {
    //            //     AppException.LogError(ex.Message, "", MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
    //            return false;
    //        }



    //    }
    //    private static bool deleteDirectoryFTP(string directoryPath, string ftpUser, string ftpPassword)
    //    {
    //        try
    //        {
    //            FtpWebRequest ftpRequest11 = (FtpWebRequest)WebRequest.Create(directoryPath);
    //            ftpRequest11.Credentials = new NetworkCredential(ftpUser, ftpPassword);
    //            ftpRequest11.Method = WebRequestMethods.Ftp.ListDirectory;
    //            FtpWebResponse response11 = (FtpWebResponse)ftpRequest11.GetResponse();
    //            StreamReader streamReader = new StreamReader(response11.GetResponseStream());
    //            string line = streamReader.ReadLine();
    //            FtpWebRequest ftpRequest1 = null;
    //            FtpWebResponse response1 = null;
    //            while (!string.IsNullOrEmpty(line))
    //            {

    //                ftpRequest1 = (FtpWebRequest)WebRequest.Create(directoryPath + "/" + line);
    //                ftpRequest1.Credentials = new NetworkCredential(ftpUser, ftpPassword);
    //                ftpRequest1.Method = WebRequestMethods.Ftp.DeleteFile;
    //                response1 = (FtpWebResponse)ftpRequest1.GetResponse();
    //                response1.Close();
    //                line = streamReader.ReadLine();
    //            }
    //            streamReader.Close();
    //            response11.Close();
    //            return true;

    //        }
    //        catch (WebException ex)
    //        {
    //            //AppException.LogError(ex.Message, "", MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
    //            return false;
    //        }

    //    }

    //    private static bool FtpDirectoryExists(string directoryPath, string ftpUser, string ftpPassword)
    //    {
    //        bool directoryExists;

    //        var request = (FtpWebRequest)WebRequest.Create(directoryPath);
    //        request.Method = WebRequestMethods.Ftp.ListDirectory;
    //        request.Credentials = new NetworkCredential(ftpUser, ftpPassword);
    //        try
    //        {
    //            using (request.GetResponse())
    //            {
    //                directoryExists = true;
    //            }


    //        }
    //        catch (WebException ex)
    //        {
    //            // AppException.LogError(ex.Message, "", MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
    //            directoryExists = false;
    //        }

    //        return directoryExists;
    //    }
    //}

}
