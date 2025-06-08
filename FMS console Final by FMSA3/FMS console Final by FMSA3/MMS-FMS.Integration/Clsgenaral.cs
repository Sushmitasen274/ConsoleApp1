using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DBLayer;
using System.Net;
using System.Data.OleDb;
using System.Data;
using System.Configuration;
namespace MMS_FMS.Integration
{
    public class Clsgenaral
    {
        //public long Get_max_no(string Col_name, string Tab_name)
        //{
        //    //long No=0;
        //    try
        //    {
        //        OleDbDataReader reader;
        //        long lngrReturn = 0;
        //        reader = Fetch("SELECT MAX(" + Col_name + ") FROM  " + Tab_name);
        //        if (reader.Read())
        //        {
        //            if (reader.GetValue(0).ToString() == "")
        //            {
        //                reader.Close();
        //                return (1);
        //            }
        //            else
        //            {

        //                lngrReturn = (Convert.ToInt64(reader.GetValue(0)) + 1);
        //                reader.Close();
        //                return lngrReturn;
        //            }
        //        }

        //        else
        //        {
        //            return (-1);
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }


        //}
        public static DataTable getmonthname(string frommonth, string tomonth)
        {
            string ConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString; 
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


                dtDetails = DBHelper.SPGetDatatable(ConString, "[proc_getmonthname]", objParam);
            }
            catch (Exception ex)
            {
                //AppException.LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), MethodBase.GetCurrentMethod().DeclaringType.ToString(), MethodBase.GetCurrentMethod().Name, "[proc_getmonthname]", "");
            }
            return dtDetails;
        }
        public static string GetLocationName(int LocationCode)
        {
            string sSql = string.Empty;
            string _locationtype = string.Empty;
            string ConString = ConfigurationManager.ConnectionStrings["ConString"].ConnectionString; try
            {
                sSql = "SELECT LM_LocName FROM [MAS_Location_Master] WHERE LM_LocCode = '" + LocationCode + "'";
                _locationtype = Convert.ToString(DBHelper.DBExecuteScalar(ConString, sSql));
            }
            catch (Exception ex)
            {
               // LogError(ex.Message, Convert.ToString(System.Web.HttpContext.Current.Session["Loccode"]), "clsGeneral", System.Reflection.MethodBase.GetCurrentMethod().Name, sSql, "");
            }
            return _locationtype;
        }

    }
}
