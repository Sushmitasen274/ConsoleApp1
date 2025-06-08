using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using DBLayer;

namespace MMS_FMS.Integration
{
    class ClsFetchannexure
    {
        public static string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;

        internal DataSet GetAnnexureDetailOneLatest(int location, int YearId, string frmonth, string tomonth)
        {
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {


                objParam[iParCount] = new OleDbParameter("@Location", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@YearId", OleDbType.VarChar);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@toMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                //AppException.LogErrorAnnexure("Before call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), //MethodBase.GetCurrentMethod().Name,Convert.ToString(LocationCode));
                if (Convert.ToString(location) == "409")
                {
                    dsDetails = DBHelper.SPGetDataSet(sConString, "rpt_GetAnnexure1ch", objParam);
                }
                else
                {
                    dsDetails = DBHelper.SPGetDataSet(sConString, "rpt_GetAnnexure1", 300, objParam);
                }
                //  AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[rpt_GetAnnexure1]", "");
            }
            return dsDetails;
        }

        public DataSet GetAnnexureDetailsTwo(int location, int YearId, string frmonth, string tomonth)
        {
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {
                objParam[iParCount] = new OleDbParameter("@Location", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@YearId", OleDbType.VarChar);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@ToMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;

                
                dsDetails = DBHelper.SPGetDataSet(ConStringNew, "rpt_GetAnnexure2", objParam);

               
            }
            catch (Exception ex)
            {
               // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[rpt_GetAnnexure2]", "");
            }
            return dsDetails;
        }

        public DataTable GetAnnexure2ADetails(int location, int YearId, string frmonth, string tomonth)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {

                objParam[iParCount] = new OleDbParameter("@locationcode", OleDbType.Integer);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromMonthID", OleDbType.Integer);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;


                objParam[iParCount] = new OleDbParameter("@ToMonthID", OleDbType.Integer);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

               // AppException.LogErrorAnnexure("Before call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
                dtDetails = DBHelper.SPGetDatatable(ConStringNew, "Proc_ledger_Annexure2a", objParam);

              //  AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
            }
            catch (Exception ex)
            {
               // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[Proc_ledger_Annexure2a]", "");
            }
            return dtDetails;
        }
        public DataTable GetAnnexureDetailsThree(int LocationCode, int YearId, string frmonth, string tomonth)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {
                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@yearid", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@frommonth", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@tomonthid", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;

               // AppException.LogErrorAnnexure("Before call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));

                dtDetails = DBHelper.SPGetDatatable(ConStringNew, "rpt_GetAnnexure3", 300, objParam);

               // AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
            }
            catch (Exception ex)
            {
               // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[rpt_GetAnnexure3]", "");
            }
            return dtDetails;
        }

        public DataTable GetAnnexureDetailsThreeuptoSep(int LocationCode, int YearId, string frmonth, string tomonth)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {
                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@yearid", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@frommonth", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@tomonthid", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;


                dtDetails = DBHelper.SPGetDatatable(ConStringNew, "rpt_GetAnnexure3_uptoSep", 300, objParam);
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[rpt_GetAnnexure3]", "");
            }
            return dtDetails;
        }
        public DataTable GetAnnexureDetailsThreeupMF(int LocationCode, int YearId, string frmonth, string tomonth)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {
                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@yearid", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@frommonth", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@tomonthid", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;


                dtDetails = DBHelper.SPGetDatatable(ConStringNew, "rpt_GetAnnexure3_upMF", 300, objParam);
            }
            catch (Exception ex)
            {
               // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[rpt_GetAnnexure3]", "");
            }
            return dtDetails;
        }

        public DataTable GetAnnexureDetailsFive(int LocationCode, int YearId, string fromMonthID, string toMonthID )
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {
                objParam[iParCount] = new OleDbParameter("@AST_Location", OleDbType.Integer);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@AST_FrmMonthId", OleDbType.Integer);
                objParam[iParCount].Value = fromMonthID;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@AST_ToMonthId", OleDbType.Integer);
                objParam[iParCount].Value = toMonthID;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@AST_YearId", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;

               // AppException.LogErrorAnnexure("Before call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));

                dtDetails = DBHelper.SPGetDatatable(ConStringNew, "rpt_Asset_Annexure5", 300, objParam);
                //AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
            }
            catch (Exception ex)
            {
               // AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[rpt_Asset_Annexure5_Test]", "");
            }
            return dtDetails;
        }
        public DataSet GetAnnexure6(int LocationCode, int YearId, string fromMonthID, string toMonthID)
        {
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["SlaveConString"].ConnectionString;
            try
            {
                objParam[iParCount] = new OleDbParameter("@UnitCode", OleDbType.VarChar);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromDate1", OleDbType.VarChar);
                objParam[iParCount].Value = fromMonthID;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@ToDate1", OleDbType.VarChar);
                objParam[iParCount].Value = toMonthID;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

               // AppException.LogErrorAnnexure("Before call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));

                dsDetails = DBHelper.SPGetDataSet(ConStringNew, "Proc_AnnexureA6", objParam);
                //AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[Proc_AnnexureA6]", "");
            }
            return dsDetails;
        }
        public DataTable getmonthname(string frommonth, string tomonth)
        {
            string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[2];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@frommonth", OleDbType.VarChar);
                objParam[iParCount].Value = frommonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@tomonth", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;


                dtDetails = DBHelper.SPGetDatatable(sConString, "[proc_getmonthname]", objParam);
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[proc_getmonthname]", "");
            }
            return dtDetails;
        }

        public  string GetLocationName(Int32 LocationCode)
        {

            string sSql = string.Empty;
            string _locationtype = string.Empty;
            string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            try
            {
                sSql = "SELECT LM_LocName +'-'+ cast(LM_LocCode as varchar) FROM [MAS_Location_Master] WHERE LM_LocCode = '" + LocationCode + "'";
                _locationtype = Convert.ToString(DBHelper.DBExecuteScalar(sConString, sSql));
            }
            catch (Exception ex)
            {
               // LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), "clsGeneral", System.Reflection.MethodBase.GetCurrentMethod().Name, sSql, "");
            }
            return _locationtype;
        }
        public  string GetAccountCodeDescription(string AccountCode, string sType = "")
        {
            string sConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            string sSql = string.Empty;
            string sResult = string.Empty;
            try
            {
                if (sType == "")
                {
                    sSql = "SELECT  COA_Description FROM [MAS_Chart_Of_Accounts] WHERE COA_FullGLCode='" + AccountCode + "'  ";
                }
                else
                {
                    sSql = "SELECT  COA_Description FROM [MAS_Chart_Of_Accounts] WHERE COA_FullGLCode='" + AccountCode + "' AND  COA_LevelCode = 'L' AND COA_Interface IN ('R','B')";
                }
                sResult = Convert.ToString(DBHelper.DBExecuteScalar(sConString, sSql));
            }
            catch (Exception ex)
            {
                //LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), "clsGeneral", System.Reflection.MethodBase.GetCurrentMethod().Name, sSql, "");
            }
            return sResult;
        }

        public DataSet GetAnnexureDetails(string sType, string sInPosAccCode = "", string sInNegAccCode = "", string sOutAccCode = "", string sUnitCode = "", string sBalType = "", string frmonth = "", string tomonth = "", string sFromDate1 = "", string sToDate1 = "")
        {
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[10];
            int iParCount = 0;
            string ConStringNew = ConfigurationManager.ConnectionStrings["SlaveConString"].ConnectionString;

            /// string ConStringNew = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString; 
            try
            {
                objParam[iParCount] = new OleDbParameter("@Type", OleDbType.VarChar);
                objParam[iParCount].Value = sType;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@InputPosAccCode", OleDbType.VarChar);
                objParam[iParCount].Value = sInPosAccCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@InputNegAccCode", OleDbType.VarChar);
                objParam[iParCount].Value = sInNegAccCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@OutputAccCode", OleDbType.VarChar);
                objParam[iParCount].Value = sOutAccCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@UnitCode", OleDbType.VarChar);
                objParam[iParCount].Value = sUnitCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@BalType", OleDbType.Char);
                objParam[iParCount].Value = sBalType;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromDate", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@ToDate", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromDate1", OleDbType.VarChar);
                objParam[iParCount].Value = sFromDate1;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@ToDate1", OleDbType.VarChar);
                objParam[iParCount].Value = sToDate1;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                ///  AppException.LogErrorAnnexure("Before call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));

                dsDetails = DBHelper.SPGetDataSet(ConStringNew, "Proc_GetFinalAnnexureDetailsA6", objParam);

                ///  AppException.LogErrorAnnexure("After call sp", MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, Convert.ToString(LocationCode));
            }
            catch (Exception ex)
            {
            }
            return dsDetails;
        }
        public DataSet GetAnnexureDetails22CorD(int LocationCode, int YearId, string frmonth, string tomonth, string fAccCode, string sAccCode, string tAccCode)
        {
            string ConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString;
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[7];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@YearId", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@From_Month", OleDbType.Integer);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@To_Month", OleDbType.Integer);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@fAccCode", OleDbType.VarChar);
                objParam[iParCount].Value = fAccCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@sAccCode", OleDbType.VarChar);
                objParam[iParCount].Value = sAccCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@tAccCode", OleDbType.VarChar);
                objParam[iParCount].Value = tAccCode;
                objParam[iParCount].Direction = ParameterDirection.Input;


                dsDetails = DBHelper.SPGetDataSet(ConString, "rpt_GetAnnexure22CorD", objParam);
                // dsDetails = DBHelper.SPGetDataSet(ConString, ProcedureName, objParam);

            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }
        internal DataSet GetAnnexureDetailTwentyTwoB(int location, int YearId, string frmonth, string tomonth)
        {
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@Location", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@ToMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;

                ERPService.ERPServiceClient objService = new ERPService.ERPServiceClient();
                if (Convert.ToString(location) == "267" || Convert.ToString(location) == "766" || Convert.ToString(location) == "252")
                {
                    dsDetails = objService.GetAnnexure22ABCircle("MMS", Convert.ToString(location), frmonth, tomonth);
                }
                else
                {
                    dsDetails = objService.GetAnnexure22AB("MMS", Convert.ToString(location), frmonth, tomonth);
                }

                //dsDetails = DBHelper.SPGetDataSet(ConString, "SP_Annexure22B", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }

        public DataSet GetAnnexureDetailTwentyTwoBinFMS(int location, int YearId, string frmonth, string tomonth)
        {
            DataSet dsDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@YearId", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@From_Month", OleDbType.Integer);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@To_Month", OleDbType.Integer);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;


                dsDetails = DBHelper.SPGetDataSet(sConString, "rpt_GetAnnexure22B", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }
        internal DataSet GetAnnexureDetailTwentyTwoA(int location, int YearId, string frmonth, string tomonth)
        {
            DataSet dsDetails = new DataSet();
           
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@Location", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = frmonth;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@ToMonthId", OleDbType.VarChar);
                objParam[iParCount].Value = tomonth;
                objParam[iParCount].Direction = ParameterDirection.Input;


                ERPService.ERPServiceClient objService = new ERPService.ERPServiceClient();

                dsDetails = objService.GetAnnexure22AB("MMS", Convert.ToString(location), frmonth, tomonth);

                //dsDetails = DBHelper.SPGetDataSet(ConString, "SP_Annexure22A", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }

        internal string Get31CB(int LocationCode, int YearId, string frmonth, string tomonth)
        {

            string sSql = string.Empty;
            string FetCB = string.Empty;
            sSql = "SELECT ISNULL(SUM(LP_ClosingBal), 0) FROM Trans_Ledger_Posting WHERE LP_MonthYearId=(SELECT  MAX(LP_MonthYearId) FROM Trans_Ledger_Posting WHERE LP_UnitCode=" + LocationCode + " AND LP_YearId="+ YearId + " AND LP_FullGLCode='31.310' AND LP_MonthYearId<='" + tomonth + "') AND LP_UnitCode=" + LocationCode + " AND LP_FullGLCode='31.310' ";
            FetCB = Convert.ToString(DBHelper.DBExecuteScalar(sConString, sSql));
            return FetCB;
        }
        public DataTable ExportAnnx19(int LocationCode, int From_Month, int To_Month, int YearId)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;
            try
            {
                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = LocationCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@YearId", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@From_Month", OleDbType.Integer);
                objParam[iParCount].Value = From_Month;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@To_Month", OleDbType.Integer);
                objParam[iParCount].Value = To_Month;
                objParam[iParCount].Direction = ParameterDirection.Input;

                dtDetails = DBHelper.SPGetDatatable(sConString, "[rpt_GetAnnexure19Final]", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, LocationCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dtDetails;
        }
        public DataTable GetAnnexure19ADetails(string locationcode, string tomonth)
        {
            DataTable dt = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[2];
            int iParCount = 0;
            try
            {
                objParam[iParCount] = new OleDbParameter("@tomonthid", OleDbType.Integer);
                objParam[iParCount].Value = Convert.ToInt32(tomonth);
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@LocationCode", OleDbType.Integer);
                objParam[iParCount].Value = Convert.ToInt32(locationcode);
                objParam[iParCount].Direction = ParameterDirection.Input;


                dt = DBHelper.SPGetDatatable(sConString, "rpt_GetAnnexure19A", objParam);

                return dt;
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, locationcode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dt;

        }
        public DataTable GetAnnexureNineteenBDetails(string sUnitCode, string FromMonthID, string ToMonthID)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@locationcode", OleDbType.Integer);
                objParam[iParCount].Value = sUnitCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@FromMonthID", OleDbType.Integer);
                objParam[iParCount].Value = FromMonthID;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;


                objParam[iParCount] = new OleDbParameter("@ToMonthID", OleDbType.Integer);
                objParam[iParCount].Value = ToMonthID;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                dtDetails = DBHelper.SPGetDatatable(sConString, "Proc_Annexure19_bbb", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, sUnitCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dtDetails;
        }
        internal DataTable GetAnnexure30(int location, int YearId, string frmonth, string tomonth)
        {
            DataTable dsDetails = new DataTable();
            DataSet dataSetDetails = new DataSet();

            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {

                objParam[iParCount] = new OleDbParameter("@FROM_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@TO_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@SM_CODE", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                //ds = DBHelper.SPGetDataSet(Connstring, "[SP_Store_Inventry_RVINV_Annexure40A]", objParam);

                ERPService.ERPServiceClient objService = new ERPService.ERPServiceClient();

                dataSetDetails = objService.GetAnnexure30("MMS", Convert.ToString(location), frmonth, tomonth);
                dsDetails = dataSetDetails.Tables[0];
                //dsDetails = DBHelper.SPGetDataSet(ConString, "SP_Annexure22B", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }

        internal DataTable GetAnnexure30A(int location, int YearId, string frmonth, string tomonth)
        {
            DataTable dsDetails = new DataTable();
            DataSet dataSetDetails = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {

                objParam[iParCount] = new OleDbParameter("@FROM_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@TO_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@SM_CODE", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                //ds = DBHelper.SPGetDataSet(Connstring, "[SP_Store_Inventry_RVINV_Annexure40A]", objParam);

                ERPService.ERPServiceClient objService = new ERPService.ERPServiceClient();

                dataSetDetails = objService.GetAnnexure30A("MMS", Convert.ToString(location), frmonth, tomonth);
                dsDetails = dataSetDetails.Tables[0];

                //dsDetails = DBHelper.SPGetDataSet(ConString, "SP_Annexure22B", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }
        internal DataTable GetAnnexure40A(int location, int YearId, string frmonth, string tomonth)
        {
            DataTable dsDetails = new DataTable();
            DataSet ds = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {

                objParam[iParCount] = new OleDbParameter("@FROM_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@TO_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@SM_CODE", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                //ds = DBHelper.SPGetDataSet(Connstring, "[SP_Store_Inventry_RVINV_Annexure40A]", objParam);

                ERPService.ERPServiceClient objService = new ERPService.ERPServiceClient();

                ds = objService.GetAnnexure40A("MMS", Convert.ToString(location), frmonth, tomonth);
                dsDetails = ds.Tables[0];
                //dsDetails=objService.GetAnnexure40A("MMS", "408", "2019-04-01", "2020-03-31");
                //dsDetails = DBHelper.SPGetDataSet(ConString, "SP_Annexure22B", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }


        internal DataTable GetAnnexure40B(int location, int YearId, string frmonth, string tomonth)
        {
            DataTable dsDetails = new DataTable();
            DataSet ds = new DataSet();
            OleDbParameter[] objParam = new OleDbParameter[3];
            int iParCount = 0;

            try
            {

                objParam[iParCount] = new OleDbParameter("@FROM_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@TO_TRANS_DATE", OleDbType.VarChar);
                objParam[iParCount].Value = "";
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@SM_CODE", OleDbType.VarChar);
                objParam[iParCount].Value = location;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                

                ERPService.ERPServiceClient objService = new ERPService.ERPServiceClient();

                ds = objService.GetAnnexure40B("MMS", Convert.ToString(location), frmonth, tomonth);
                dsDetails = ds.Tables[0]; 
                }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, location.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dsDetails;
        }
        public DataTable GetAnnexureThirtyOne(int sUnitCode, int YearId, int From_Month, int To_Month)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[4];
            int iParCount = 0;

            try
            {


                objParam[iParCount] = new OleDbParameter("@locationcode", OleDbType.Integer);
                objParam[iParCount].Value = sUnitCode;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@YearId", OleDbType.Integer);
                objParam[iParCount].Value = YearId;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@From_Month", OleDbType.Integer);
                objParam[iParCount].Value = From_Month;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@To_Month", OleDbType.Integer);
                objParam[iParCount].Value = To_Month;
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                dtDetails = DBHelper.SPGetDatatable(sConString, "rpt_GetAnnexure31", objParam);
            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, sUnitCode.ToString(), System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dtDetails;
        }
        public DataTable GetDateFormMonth(DataTable dtmonths)
        {
            DataTable dtDetails = new DataTable();
            OleDbParameter[] objParam = new OleDbParameter[2];
            string[] DateResult = new string[4];
            int iParCount = 0;

            try
            {
                objParam[iParCount] = new OleDbParameter("@frommonth", OleDbType.VarChar);
                objParam[iParCount].Value = dtmonths.Rows[0]["YMC_Month_Name"];
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;

                objParam[iParCount] = new OleDbParameter("@tomonth", OleDbType.VarChar);
                objParam[iParCount].Value = dtmonths.Rows[1]["YMC_Month_Name"];
                objParam[iParCount].Direction = ParameterDirection.Input;
                iParCount = iParCount + 1;


                dtDetails = DBHelper.SPGetDatatable(sConString, "[proc_getmonthDates]", objParam);

            }
            catch (Exception ex)
            {
                AppException.WritetoFile(ex.Message, "", System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString(), System.Reflection.MethodBase.GetCurrentMethod().Name, "", "");
            }
            return dtDetails;
        }


    }
}
