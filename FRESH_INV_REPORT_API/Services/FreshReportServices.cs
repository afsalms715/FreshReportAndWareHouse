using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using FRESH_INV_REPORT_API.Models;
using System.Globalization;

namespace FRESH_INV_REPORT_API.Services
{
    public class FreshReportServices
    {
        public string connectionString;
        public FreshReportServices(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public string GetFormatedDate(string date)
        {
            DateTime FormatedDate = DateTime.Parse(date);
            return FormatedDate.ToString("dd-MMM-yy").ToUpper();
        }

        //GET FOR OPENING ID OF INVENTORY
        public List<OpeningInvDtl> GetOpeningInvId(string fromDate,string loc)
        {
            List<OpeningInvDtl> ListOfopeningInvDtl = new List<OpeningInvDtl>();
            using(OracleConnection connection=new OracleConnection(connectionString))
            {
                string FormatedFromDate = GetFormatedDate(fromDate);
                string query = @"SELECT DISTINCT INV_NO,
                                    CASE
                                        WHEN SUBSTR(HIERARCHY, 1, 9) = '1_1S_1S03' THEN 'FRESH'
                                    END AS CATEGORY
                                FROM GOLDPROD.GRAND_INV_DETL_02@GOLD_SERVER
                                WHERE INV_DATE ='"+ FormatedFromDate + @"' AND INV_SITE ="+loc+@"
                                AND SUBSTR(HIERARCHY, 1, 9) = '1_1S_1S03'";
                connection.Open();
                using(OracleCommand cmd=new OracleCommand(query, connection))
                {
                    using(OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                OpeningInvDtl openingInvDtl = new OpeningInvDtl();
                                openingInvDtl.OpeningInvId = int.Parse(reader["INV_NO"].ToString());
                                openingInvDtl.Category = reader["CATEGORY"].ToString();
                                ListOfopeningInvDtl.Add(openingInvDtl);
                            }
                        }
                    }
                }
                connection.Close();
            }
            return ListOfopeningInvDtl;
        }

        //GET IVISION INVENTRY ID AND CLOSING INV ID
        public List<ClosingDtl> GetClosingInvDtl(string ToDate,string LocCode)
        {
            List<ClosingDtl> ListOfclosingDtl = new List<ClosingDtl>();
            using (OracleConnection connection =new OracleConnection(connectionString))
            {
                string FormatedToDate = GetFormatedDate(ToDate);
                string query = @"SELECT 
                                    DISTINCT INV_NO CLOSE_INV_ID,
                                    CASE
                                        WHEN SUBSTR(HIERARCHY, 1, 9) = '1_1S_1S03' THEN 'FRESH'
                                    END AS CATEGORY
                                FROM GOLDPROD.GRAND_INV_DETL_02@GOLD_SERVER
                                WHERE INV_DATE ='" + FormatedToDate + @"' 
                                AND SUBSTR(HIERARCHY, 1, 9) = '1_1S_1S03'
                                AND INV_SITE = " + LocCode + @"";
                connection.Open();
                using(OracleCommand cmd=new OracleCommand(query,connection))
                {
                    using(OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                ClosingDtl closingDtl = new ClosingDtl();
                                closingDtl.ClosingInvId =int.Parse(reader["CLOSE_INV_ID"].ToString());
                                closingDtl.Category = reader["CATEGORY"].ToString();
                                ListOfclosingDtl.Add(closingDtl);
                            }
                        }
                    }
                }
                connection.Close();
            }
            return ListOfclosingDtl;
        }

        //get Invento Inventory Details
        public InventoInvDtl GetInventoInvDtl(string ToDate, string LocCode)
        {
            InventoInvDtl inventoInvDtl = new InventoInvDtl();
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                string FormatedToDate = GetFormatedDate(ToDate);
                string query = @"SELECT DISTINCT INVH_NO,INVH_INV_NAME FROM IMAN_OFFLINE_INV_HEAD WHERE INVH_INV_DATE='"+FormatedToDate+@"' AND INVH_SITE="+LocCode+@"";
                connection.Open();
                using (OracleCommand cmd = new OracleCommand(query, connection))
                {
                    using (OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                inventoInvDtl.InventoInvId = int.Parse(reader["INVH_NO"].ToString());
                                inventoInvDtl.InventoInvName = reader["INVH_INV_NAME"].ToString();
                            }
                        }
                    }
                }
                connection.Close();
            }
            return inventoInvDtl;
        }
    }
}
