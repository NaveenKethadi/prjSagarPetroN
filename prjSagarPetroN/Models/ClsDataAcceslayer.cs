using Focus.Common.DataStructs;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace prjSagarPetroN.Models
{
    public class ClsDataAcceslayer
    {
        string error = "";
        public int GetQueryExe(string strSelQry, int CompId, ref string error)
        {
            try
            {
                //  DataSet ds;

                Database obj = Focus.DatabaseFactory.DatabaseWrapper.GetDatabase(CompId);
                System.Data.Common.DbConnection dbcon = obj.CreateConnection();
                //dbcon.ConnectionString = obj.ConnectionString;
                System.Data.Common.DbCommand cmd = dbcon.CreateCommand();

                // return (obj.ExecuteDataSet(CommandType.Text, strSelQry));


                cmd.CommandTimeout = 0;
                cmd.CommandText = strSelQry;
                int Result = obj.ExecuteNonQuery(cmd);
                cmd.Dispose();
                dbcon.Dispose();
                return Result;
            }
            catch (Exception e)
            {
                error = e.Message;
                return 0;
            }
        }
        public static int GetDateToInt(DateTime dt)
        {
            int val;
            val = Convert.ToInt16(dt.Year) * 65536 + Convert.ToInt16(dt.Month) * 256 + Convert.ToInt16(dt.Day);
            return val;
        }

        public static Date GetIntToDate(int iDate)
        {
            try
            {
                return (new Date(iDate, CalendarType.Gregorean));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int GetVouchertype(string InvName, int CompId)
        {
            string strsql = "Select iVouchertype,sName,sabbr from ccore_vouchers_0 where sabbr='" + InvName + "'";
            DataSet Dst2 = ClsDataAcceslayer.GetData1(strsql, CompId, ref error);

            if (Dst2 != null && Dst2.Tables.Count > 0 && Dst2.Tables[0].Rows.Count > 0)
            {
                return Convert.ToInt32(Dst2.Tables[0].Rows[0]["iVouchertype"]);
            }
            else
            {
                return 0;
            }
        }

        public static DataSet GetData1(string strSelQry, int CompId, ref string error)
        {
            try
            {
                ClsDataAcceslayer.writeLog(strSelQry);
                DataSet ds;
                Database obj = Focus.DatabaseFactory.DatabaseWrapper.GetDatabase(CompId);
                System.Data.Common.DbConnection dbcon = obj.CreateConnection();
                //dbcon.ConnectionString = obj.ConnectionString;
                System.Data.Common.DbCommand cmd = dbcon.CreateCommand();

                // return (obj.ExecuteDataSet(CommandType.Text, strSelQry));


                cmd.CommandTimeout = 0;
                cmd.CommandText = strSelQry;
                ds = obj.ExecuteDataSet(cmd);
                cmd.Dispose();
                dbcon.Dispose();
                return ds;
            }
            catch (Exception e)
            {
                error = e.Message;
                return null;
            }
        }

        internal static DataSet GetData1(string retrievequery, int cid, ref object error)
        {
            throw new NotImplementedException();
        }

        public static void writeLog(string content)
        {

            StreamWriter objSw = null;

            try
            {
                string sFilePath = @"C:\Windows\Temp\" + DateTime.Now.ToString("ddMM") + "prjSagarPetroN.txt";
                objSw = new StreamWriter(sFilePath, true);
                objSw.WriteLine(DateTime.Now.ToString() + " " + content + Environment.NewLine);
            }
            catch (Exception )
            {
                
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }
        }

        public void SetLog(string content)
        {
            StreamWriter objSw = null;
            try
            {
                string sFilePath = System.IO.Path.GetTempPath() + "prjSagarPetroN.log" + DateTime.Now.Date.ToString("ddMMyyyy") + ".txt";
                objSw = new StreamWriter(sFilePath, true);
                objSw.WriteLine(DateTime.Now.ToString() + " " + content + Environment.NewLine);
            }
            catch (Exception ex)
            {
                //SetLog("Error -" + ex.Message);
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }
        }


        public int GetExecute(string strSelQry, int CompId, ref string error)
        {
            error = "";
            try
            {
                Database obj = Focus.DatabaseFactory.DatabaseWrapper.GetDatabase(CompId);
                return (obj.ExecuteNonQuery(CommandType.Text, strSelQry));
            }
            catch (Exception e)
            {
                error = e.Message;
                FConvert.LogFile("prjSagarPetroN.log", DateTime.Now.ToString() + " GetExecute :" + error + "---" + strSelQry);
                return 0;
            }
        }
    }
}