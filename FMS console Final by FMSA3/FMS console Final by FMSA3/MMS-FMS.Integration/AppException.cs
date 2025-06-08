
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Web;
using System.Diagnostics;
using System.Configuration;
using System;
using DBLayer;
using System.Data.OleDb;
using System.Threading.Tasks;

using System.Data;
using System.Reflection;

public interface IAppException
{
    void LogError(string Message, string UnitID, string MyClassName, string MyFunctionName, string sQuery, string sSource);
}


public class AppException
{
    static string FilePath = null;


    //public static void LogError(string Message, string UnitID, string MyClassName, string MyFunctionName, string sQuery, string sSource="F")
    //{
    //    CustOledbConnection objCon = new CustOledbConnection("idocs@123");
    //    string ErrMsg = string.Empty;
    //    //FilePath = System.Web.HttpContext.Current.Server.MapPath("~/ErrLog.txt");
    //    if(Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["ErrorLog"])=="ON")
    //    {
    //        long strElid = objCon.Get_max_no("EL_ID", "TBLERRORLOG");
    //       // string LocUserName = Convert.ToString(System.Web.HttpContext.Current.Session["LocationCodeName"]) + " - " + Convert.ToString(System.Web.HttpContext.Current.Session["UserLogInName"]);
    //       string strQry = "INSERT INTO TBLERRORLOG(EL_ID,EL_PAGENAME,EL_EVENTNAME,EL_ERRORMESSAGE,EL_STACKTRACE,EL_ENTRYDATE, EL_RECORD_BY,EL_Loc_user_Name) ";
    //       strQry += " VALUES('" + objCon.Get_max_no("EL_ID", "TBLERRORLOG") + "','" + MyClassName + "','" + MyFunctionName + "',";
    //       strQry += "'" + Message.Trim().Replace("'", "`") + "','" + sQuery.Trim().Replace("'", "`") + "',GETDATE(),'','')";

    //        objCon.Execute(strQry);
    //        return;
    //    }
    //    FilePath = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["ErrorLogPath"]);



    //    string FileName = FilePath + "/" + System.DateTime.Now.ToString("dd.MM.yyyy") + "_" + "ErrLog.txt";

    //    if(!File.Exists(FileName))
    //    {
    //        //File.Create(FileName);
    //        File.Create(FileName).Close();

    //    }

    //    // Calculate GMT offset
    //    int GmtOffset = DateTime.Compare(DateTime.Now, DateTime.UtcNow);
    //    string GmtPrefix = null;
    //    if (GmtOffset > 0)
    //    {
    //        GmtPrefix = "+";
    //    }
    //    else
    //    {
    //        GmtPrefix = "";
    //    }
    //    // Create DateTime string
    //    string ErrorDateTime = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString() + " @ " + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":" + DateTime.Now.Second.ToString() + " (GMT " + GmtPrefix + GmtOffset.ToString() + ")";
    //    // Write message to error file
    //    try
    //    {


    //            StreamWriter MsStreamWriter = new StreamWriter(FileName, true);
    //            MsStreamWriter.WriteLine(Environment.NewLine);
    //            MsStreamWriter.WriteLine("Date And Time - " + ErrorDateTime);
    //            MsStreamWriter.WriteLine("Unit ID - " + UnitID);
    //            MsStreamWriter.WriteLine("Class Name - " + MyClassName);
    //            MsStreamWriter.WriteLine("Function Name - " + MyFunctionName);
    //            MsStreamWriter.WriteLine("SQL Query - " + sQuery);
    //            MsStreamWriter.WriteLine("Error Message - " + Message);
    //            MsStreamWriter.WriteLine("##################################################################");
    //            MsStreamWriter.Close();
    //            ErrMsg = Message + " in function '" + MyFunctionName + "' of class file '" + MyClassName + "' " + Environment.NewLine + Environment.NewLine + "SQL Query is " + sQuery + "";
    //            if (sSource == "C")
    //            {
    //                //SendEmail(ErrMsg, UnitID);
    //            }
    //            else if (sSource == "F")
    //            {
    //                //Don't Send Mail
    //                //SendEmail(ErrMsg, UnitID)
    //            }//update UI

    //            SendMailOnWOSyncExcp(ErrMsg);
    //    }
    //    catch (Exception ex)
    //    {
    //        throw;
    //    }
    //}

    public static void SendMailOnWOSyncExcp(string sMsg)
    {
        try
        {
            if (Convert.ToString(ConfigurationManager.AppSettings["SendEmailForWOSyncException"]).ToUpper().Equals("ON"))
            {
                string ToMailId = Convert.ToString(ConfigurationManager.AppSettings["WOSyncEmailTo"]);
                string CCEmailId = Convert.ToString(ConfigurationManager.AppSettings["WOSyncCCTo"]);

                MailMessage mail = new MailMessage();
                mail.To.Add(ToMailId);
                mail.From = new MailAddress("support@ideainfinityit.com", "Financial Management System");
                if (CCEmailId != "")
                {
                    mail.CC.Add(CCEmailId);
                }
                mail.Subject = "CESCOM-FAMS Application - Alert";
                mail.IsBodyHtml = true;
                mail.Body = "Hi Sir/Madam,  " + Environment.NewLine + "Error - while generating .csv files of CR completed work orders." + Environment.NewLine + sMsg + Environment.NewLine + "Rgds" + Environment.NewLine + "Team FAMS";
                mail.BodyEncoding = System.Text.Encoding.UTF8;

                SmtpClient smtp = new SmtpClient("smtp.bizmail.yahoo.com", 25);
                smtp.Credentials = new System.Net.NetworkCredential
                     ("support@ideainfinityit.com", "kvuqroyceqiqneav");

                smtp.EnableSsl = true;
                smtp.Send(mail);
            }
        }
        catch (Exception ex)
        {
            throw;
        }
    }
    public static void WritetoFile(string Message, string UnitID, string MyClassName, string MyFunctionName, string sQuery, string sSource = "F")
    {
        try
        {
            string sFolderPath = Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["AnnexureLog"]) + "AnnexureLog/" + DateTime.Now.ToString("yyyyMM");
            if (Convert.ToString(System.Configuration.ConfigurationManager.AppSettings["Annexurelogconfig"]) == "OFF")
            {
                return;
            }
            else
            {
                if (!Directory.Exists(sFolderPath))
                {
                    Directory.CreateDirectory(sFolderPath);

                }
                string sPath = sFolderPath + "//" + DateTime.Now.ToString("yyyyMMdd") + "-Annexurelog.csv";
                string annexuretype = "Annexure_22series";
                File.AppendAllText(sPath, "  Annexure-type : " + annexuretype + ",  UnitID : " + UnitID + ", ClassName:" + MyClassName + ",FunctionName:" + MyFunctionName + ",logDate :" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + ", Message --" + Message + Environment.NewLine);
            }
        }
        catch (Exception ex)
        {
            return;
        }
    }

}

