using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using FRESH_INV_REPORT_API.Models;
using System.Globalization;
using OfficeOpenXml;
using System.IO;

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

        //GET CLOSING INV ID
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
        public List<InventoInvDtl> GetInventoInvDtl(string ToDate, string LocCode)
        {
            
            List<InventoInvDtl> inventoInvDtls = new List<InventoInvDtl>();
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
                                InventoInvDtl inventoInvDtl = new InventoInvDtl();
                                inventoInvDtl.InventoInvId = int.Parse(reader["INVH_NO"].ToString());
                                inventoInvDtl.InventoInvName = reader["INVH_INV_NAME"].ToString();
                                inventoInvDtls.Add(inventoInvDtl);
                            }
                        }
                    }
                }
                connection.Close();
            }
            return inventoInvDtls;
        }

        //testing generation excel
        public void ExcelGeneration()
        {
            List<ExcelDemo> ExcelDatas = new List<ExcelDemo>();
            using(OracleConnection connection =new OracleConnection(connectionString))
            {
                string query = @"SELECT LOCATION_ID, SECTION_CODE, SECTION_NAME, SALE FROM AFSAL_TEMP";
                connection.Open();
                using(OracleCommand cmd=new OracleCommand(query, connection))
                {
                    using(OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                ExcelDemo excelRow = new ExcelDemo();
                                excelRow.LocationId = int.Parse(reader["LOCATION_ID"].ToString());
                                excelRow.SectionCode = reader["SECTION_CODE"].ToString();
                                excelRow.SectionName = reader["SECTION_NAME"].ToString();
                                excelRow.Sale = reader["SALE"].ToString();
                                ExcelDatas.Add(excelRow);
                            }
                        }
                    }
                }
                connection.Close();
                string CurrentDirectry = Environment.CurrentDirectory;
                string filePath = Path.Combine(CurrentDirectry, "ReportFiles", "ExcelReport" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                SaveExcelFile(ExcelDatas, filePath);
            }
        }

        //saving file
        public void SaveExcelFile(List<ExcelDemo> ExcelDatas,string FilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package= new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells[1, 1].Value = "LocationId";
                worksheet.Cells[1, 2].Value = "SectionCode";
                worksheet.Cells[1, 3].Value = "SectionName";
                worksheet.Cells[1, 4].Value = "Sale";
                worksheet.Cells[1, 5].Value = "sale+location";
                int row = 2;
                foreach(var data in ExcelDatas)
                {
                    worksheet.Cells[row, 1].Value = data.LocationId;
                    worksheet.Cells[row, 2].Value = data.SectionCode;
                    worksheet.Cells[row, 3].Value = data.SectionName;
                    worksheet.Cells[row, 4].Value = data.Sale;
                    worksheet.Cells["E"+row.ToString()].Formula ="A"+ row.ToString()+"&"+ "B" + row.ToString();
                    row++;
                }
                File.WriteAllBytes(FilePath, package.GetAsByteArray());
            }
        }
    }
}
