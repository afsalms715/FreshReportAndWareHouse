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

        public string GetFormatedDate(string date,bool isFromDate=false)
        {
            DateTime FormatedDate = DateTime.Parse(date);
            if (isFromDate)
            {
                FormatedDate.AddDays(1);
            }
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
        public string ExcelGeneration(string query,string location,string toDate)
        {
            List<ExcelDemo> ExcelDatas = new List<ExcelDemo>();
            using(OracleConnection connection =new OracleConnection(connectionString))
            {
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
                                excelRow.PRODUCT_CODE = int.Parse(reader["PRODUCT_CODE"].ToString());
                                excelRow.SU =int.Parse(reader["SU"].ToString());
                                excelRow.BARCODE = long.Parse(reader["BARCODE"].ToString());
                                excelRow.STOCK_UNIT = reader["STOCK_UNIT"].ToString();
                                excelRow.SECTION = reader["SECTION"].ToString();
                                excelRow.SU_DESC = reader["SU_DESC"].ToString();
                                excelRow.UNIT_COST = float.Parse(reader["UNIT_COST"].ToString());
                                excelRow.OPENING_QTY = float.Parse(reader["OPENING_QTY"].ToString());
                                excelRow.OPENING_VAL = float.Parse(reader["OPENING_VAL"].ToString());
                                excelRow.NET_PURCHASE_QTY = float.Parse(reader["NET_PURCHASE_QTY"].ToString());
                                excelRow.NET_PURCHASE_VAL = float.Parse(reader["NET_PURCHASE_VAL"].ToString());
                                excelRow.SALES_QTY = float.Parse(reader["SALES_QTY"].ToString());
                                excelRow.SALES_VAL = float.Parse(reader["SALES_VAL"].ToString());
                                excelRow.WASTAGE_QTY = float.Parse(reader["WASTAGE_QTY"].ToString());
                                excelRow.WASTAGE_VAL = float.Parse(reader["WASTAGE_VAL"].ToString());
                                excelRow.SYS_QTY = float.Parse(reader["SYS_QTY"].ToString());
                                excelRow.SYS_VAL = float.Parse(reader["SYS_VAL"].ToString());
                                excelRow.CLOSING_QTY = float.Parse(reader["CLOSING_QTY"].ToString());
                                excelRow.CLOSING_VAL = float.Parse(reader["CLOSING_VAL"].ToString());
                                excelRow.UNKNWON_WASTAGE_QTY = float.Parse(reader["UNKNWON_WASTAGE_QTY"].ToString());
                                excelRow.UNKNOWN_WASTAGE_VAL = float.Parse(reader["UNKNOWN_WASTAGE_VAL"].ToString());
                                excelRow.COST_OF_SALE = float.Parse(reader["COST_OF_SALE"].ToString());
                                excelRow.PROFIT = float.Parse(reader["PROFIT"].ToString());
                                excelRow.GP = float.Parse(reader["GP"].ToString());
                                ExcelDatas.Add(excelRow);
                            }
                        }
                    }
                }
                connection.Close();
                string CurrentDirectry = Environment.CurrentDirectory;
                string filePath = Path.Combine(CurrentDirectry, "ReportFiles", "FRESH_GP_REPORT_"+location+"_"+toDate+".xlsx");
                SaveExcelFile(ExcelDatas, filePath);
                return filePath;
            }
        }

        public string GenerateFreshReport(string Location, string OpenInvId, string CloseInvId, string InventoInvId, string Sections, string FromDate, string ToDate)
        {
            string query = @"SELECT
                                PRODUCT_CODE,
                                SU,BARCODE,
                                STOCK_UNIT,
                                SECTION,
                                SU_DESC,
                                NVL(UNIT_COST,0) UNIT_COST,
                                NVL(OPENING_QTY,0) OPENING_QTY,
                                NVL(OPENING_VAL,0) OPENING_VAL,
                                NVL(NET_PURCHASE_QTY,0) NET_PURCHASE_QTY,
                                NVL(NET_PURCHASE_VAL,0) NET_PURCHASE_VAL,
                                NVL(SALES_QTY,0) SALES_QTY,
                                NVL(SALES_VAL,0) SALES_VAL,
                                NVL(WASTAGE_QTY,0) WASTAGE_QTY,
                                NVL(WASTAGE_VAL,0) WASTAGE_VAL,
                                NVL((NVL(OPENING_QTY,0)+NVL(NET_PURCHASE_QTY,0)-NVL(SALES_QTY,0)-NVL(WASTAGE_QTY,0)),0) AS SYS_QTY,
                                NVL((NVL(OPENING_QTY,0)+NVL(NET_PURCHASE_QTY,0)-NVL(SALES_QTY,0)-NVL(WASTAGE_QTY,0))*UNIT_COST,0) AS SYS_VAL,
                                NVL(CLOSING_QTY,0) CLOSING_QTY,
                                NVL(CLOSING_QTY*UNIT_COST,0) CLOSING_VAL,
                                NVL((NVL(OPENING_QTY,0)+NVL(NET_PURCHASE_QTY,0)-NVL(SALES_QTY,0)-NVL(WASTAGE_QTY,0))-NVL(CLOSING_QTY,0),0) UNKNWON_WASTAGE_QTY,
                                NVL(((NVL(OPENING_QTY,0)+NVL(NET_PURCHASE_QTY,0)-NVL(SALES_QTY,0)-NVL(WASTAGE_QTY,0))-NVL(CLOSING_QTY,0))*UNIT_COST,0) UNKNOWN_WASTAGE_VAL,
                                NVL((NVL(OPENING_VAL,0)+NVL(NET_PURCHASE_VAL,0)-NVL(CLOSING_VAL,0)),2) COST_OF_SALE,
                                NVL(SALES_VAL,0)-NVL((NVL(OPENING_VAL,0)+NVL(NET_PURCHASE_VAL,0)-NVL(CLOSING_VAL,0)),2) PROFIT,
                                (CASE WHEN NVL(SALES_VAL,0)=0 THEN 0 ELSE (
                                ROUND(((NVL(SALES_VAL,0)-NVL((NVL(OPENING_VAL,0)+NVL(NET_PURCHASE_VAL,0)-NVL(CLOSING_VAL,0)),2))/NVL(SALES_VAL,0))*100,2)) END) GP
                            FROM 
                                (
                                SELECT A.*,
                                --
                                    (SELECT STOCK_UNIT FROM GRAND_PRD_MASTER_FULL WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU) STOCK_UNIT,
                                    (SELECT SECTION_NAME FROM GRAND_PRD_MASTER_FULL WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU) SECTION,BARCODE BAR,
                                    --
                                    GET_SU_DESC(PRODUCT_CODE,SU) SU_DESC,
                                    --
                                    (SELECT MAX(NVL(PHY,0)) FROM GOLDPROD.GRAND_INV_DETL_01@GOLD_SERVER
                                    WHERE INV_SITE="+ Location + @" AND INV_NO="+ OpenInvId + @" AND ART_CODE=A.PRODUCT_CODE AND ART_SU=A.SU) OPENING_QTY,
                                    --
                                    (SELECT MAX(NVL(PHY,0)*NVL(STOCK_PP,0)) FROM GOLDPROD.GRAND_INV_DETL_01@GOLD_SERVER
                                    WHERE INV_SITE=" + Location + @" AND INV_NO=" + OpenInvId + @" AND ART_CODE=A.PRODUCT_CODE AND ART_SU=A.SU) OPENING_VAL,
                                    --
                                    (SELECT SUM(NVL(QTY_WEIGHT,0)) FROM KPI_STK_MVT@ZEDEYE WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU
                                    AND MVT_DATE BETWEEN '"+ GetFormatedDate(FromDate,true) + @"' AND '"+ GetFormatedDate(ToDate) + @"' AND MVT_CODE IN (1,25,50,51,175,176,200,201,2,75,225) AND LOCATION_ID=" + Location + @") NET_PURCHASE_QTY,
                                    --
                                    (SELECT SUM(NVL(COST_VALUE_BEF_MK,0)) FROM KPI_STK_MVT@ZEDEYE WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU
                                    AND MVT_DATE BETWEEN '" + GetFormatedDate(FromDate, true) + @"' AND '" + GetFormatedDate(ToDate) + @"' AND MVT_CODE IN (1,25,50,51,175,176,200,201,2,75,225) AND LOCATION_ID=" + Location + @") NET_PURCHASE_VAL,
                                    --
                                    (SELECT SUM(NVL(QTY_WEIGHT,0)) *-1 FROM KPI_STK_MVT@ZEDEYE WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU
                                    AND MVT_DATE BETWEEN '" + GetFormatedDate(FromDate, true) + @"' AND '" + GetFormatedDate(ToDate) + @"' AND MVT_CODE IN (150,152) AND LOCATION_ID=" + Location + @") SALES_QTY,
                                    --
                                    (SELECT SUM(NVL(SALES_VALUE_W_VAT,0)-NVL(VAT_VALUE,0)) *-1 FROM KPI_STK_MVT@ZEDEYE WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU
                                    AND MVT_DATE BETWEEN '" + GetFormatedDate(FromDate, true) + @"' AND '" + GetFormatedDate(ToDate) + @"' AND MVT_CODE IN (150,152) AND LOCATION_ID=" + Location + @") SALES_VAL,
                                    --
                                    (SELECT SUM(NVL(QTY_WEIGHT,0)) *-1 FROM KPI_STK_MVT@ZEDEYE WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU
                                    AND MVT_DATE BETWEEN '" + GetFormatedDate(FromDate, true) + @"' AND '" + GetFormatedDate(ToDate) + @"' AND MVT_CODE IN (101) AND LOCATION_ID=" + Location + @") WASTAGE_QTY,
                                    --
                                    (SELECT SUM(NVL(COST_VALUE_BEF_MK,0)) *-1 FROM KPI_STK_MVT@ZEDEYE WHERE PRODUCT_CODE=A.PRODUCT_CODE AND SU=A.SU
                                    AND MVT_DATE BETWEEN '" + GetFormatedDate(FromDate, true) + @"' AND '" + GetFormatedDate(ToDate) + @"' AND MVT_CODE IN (101) AND LOCATION_ID=" + Location + @") WASTAGE_VAL,
                                    --
                                    (SELECT 
                                    SUM(NVL(INV_PHY_STOCK,0)*NVL(CONV,1))PHY
                                        FROM IMAN_OFFLINE_INV WHERE INV_HID="+ InventoInvId + @" AND INV_GOLD_CODE=A.PRODUCT_CODE AND INV_STOCK_SU=A.SU) CLOSING_QTY,
                                    ------
                                    (SELECT MAX(NVL(PHY,0)*NVL(STOCK_PP,0)) FROM GOLDPROD.GRAND_INV_DETL_01@GOLD_SERVER
                                    WHERE INV_SITE=" + Location + @" AND INV_NO="+ CloseInvId + @" AND ART_CODE=A.PRODUCT_CODE AND ART_SU=A.SU) CLOSING_VAL,
                                    --
                                    (SELECT MAX(NVL(STOCK_PP,0)) FROM GOLDPROD.GRAND_INV_DETL_01@GOLD_SERVER
                                    WHERE INV_SITE=" + Location + @" AND INV_NO=" + CloseInvId + @" AND ART_CODE=A.PRODUCT_CODE AND ART_SU=A.SU) UNIT_COST
                                FROM  
                                    ASQ_TEMP_DXB_FRESH_MASTER_01 A 
                                WHERE 
                                    SECTION_CODE IN("+ Sections +@")
                                ) 
                            WHERE 
                                (NVL(OPENING_QTY,0) <> 0 OR NVL(NET_PURCHASE_QTY,0) <> 0 OR NVL(SALES_QTY,0) <> 0 OR NVL(WASTAGE_QTY,0) <> 0 OR NVL(CLOSING_QTY,0) <> 0)";
            
            return ExcelGeneration(query, Location, GetFormatedDate(ToDate));
        }

        //saving file
        public void SaveExcelFile(List<ExcelDemo> ExcelDatas,string FilePath)
        {
            //HOT FOOD INDUSTRIAL --HOT FOOD
            //FRUITS  AND  VEGETABLES(W)
            //FISHERY (W)
            //BUTCHERY (W)
            //BAKERY INHOUSE
            //DELICATTESEN (W)
            List<ExcelDemo> Hot_food= new List<ExcelDemo>();
            List<ExcelDemo> Fruits_And_veg = new List<ExcelDemo>();
            List<ExcelDemo> Fish = new List<ExcelDemo>();
            List<ExcelDemo> Butchery = new List<ExcelDemo>();
            List<ExcelDemo> Bakery = new List<ExcelDemo>();
            List<ExcelDemo> Deli = new List<ExcelDemo>();
            //split as sections
            foreach (var data in ExcelDatas)
            { 
                if(data.SECTION== "HOT FOOD INDUSTRIAL")
                {
                   Hot_food.Add(data);
                }
                if (data.SECTION == "FRUITS  AND  VEGETABLES(W)")
                {
                    Fruits_And_veg.Add(data);
                }
                if (data.SECTION == "FISHERY (W)")
                {
                    Fish.Add(data);
                }
                if (data.SECTION == "BUTCHERY (W)")
                {
                    Butchery.Add(data);
                }
                if (data.SECTION == "BAKERY INHOUSE")
                {
                    Bakery.Add(data);
                }
                if(data.SECTION == "DELICATTESEN (W)")
                {
                    Deli.Add(data);
                }
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                if (Hot_food.Count != 0)
                {
                    var worksheet = package.Workbook.Worksheets.Add("HOT FOOD INDUSTRIAL");
                    SheetCreation(worksheet, Hot_food);
                }
                if (Fruits_And_veg.Count != 0)
                {
                    var worksheet = package.Workbook.Worksheets.Add("FRUITS  AND  VEGETABLES(W)");
                    SheetCreation(worksheet, Fruits_And_veg);
                }
                if (Fish.Count != 0)
                {
                    var worksheet = package.Workbook.Worksheets.Add("FISHERY (W)");
                    SheetCreation(worksheet, Fish);
                }
                if (Butchery.Count != 0)
                {
                    var worksheet = package.Workbook.Worksheets.Add("BUTCHERY (W)");
                    SheetCreation(worksheet, Butchery);
                }
                if (Bakery.Count != 0)
                {
                    var worksheet = package.Workbook.Worksheets.Add("BAKERY INHOUSE");
                    SheetCreation(worksheet, Bakery);
                }
                if (Deli.Count != 0)
                {
                    var worksheet = package.Workbook.Worksheets.Add("DELICATTESEN (W)");
                    SheetCreation(worksheet, Deli);
                }
                File.WriteAllBytes(FilePath, package.GetAsByteArray());
            }
        }

        public void SheetCreation(ExcelWorksheet worksheet, List<ExcelDemo> sheetData)
        {
            
                worksheet.Cells[1, 1].Value = "PRODUCT_CODE";
                worksheet.Cells[1, 2].Value = "SU";
                worksheet.Cells[1, 3].Value = "STOCK_UNIT";
                worksheet.Cells[1, 4].Value = "BARCODE";
                worksheet.Cells[1, 5].Value = "SECTION";
                worksheet.Cells[1, 6].Value = "SU_DESC";
                worksheet.Cells[1, 7].Value = "UNIT_COST";
                worksheet.Cells[1, 8].Value = "OPENING_QTY";
                worksheet.Cells[1, 9].Value = "OPENING_VAL";
                worksheet.Cells[1, 10].Value = "NET_PURCHASE_QTY";
                worksheet.Cells[1, 11].Value = "NET_PURCHASE_VAL";
                worksheet.Cells[1, 12].Value = "SALES_QTY";
                worksheet.Cells[1, 13].Value = "SALES_VAL";
                worksheet.Cells[1, 14].Value = "WASTAGE_QTY";
                worksheet.Cells[1, 15].Value = "WASTAGE_VAL";
                worksheet.Cells[1, 16].Value = "SYS_QTY";
                worksheet.Cells[1, 17].Value = "SYS_VAL";
                worksheet.Cells[1, 18].Value = "CLOSING_QTY";
                worksheet.Cells[1, 19].Value = "CLOSING_VAL";
                worksheet.Cells[1, 20].Value = "UNKNWON_WASTAGE_QTY";
                worksheet.Cells[1, 21].Value = "UNKNOWN_WASTAGE_VAL";
                worksheet.Cells[1, 22].Value = "COST_OF_SALE";
                worksheet.Cells[1, 23].Value = "PROFIT";
                worksheet.Cells[1, 24].Value = "GP";
                int row = 2;
                foreach (var data in sheetData)
                {
                    worksheet.Cells[row, 1].Value = data.PRODUCT_CODE;
                    worksheet.Cells[row, 2].Value = data.SU;
                    worksheet.Cells[row, 3].Value = data.STOCK_UNIT;
                    worksheet.Cells[row, 4].Value = data.BARCODE;
                    worksheet.Cells[row, 5].Value = data.SECTION;
                    worksheet.Cells[row, 6].Value = data.SU_DESC;
                    worksheet.Cells[row, 7].Value = data.UNIT_COST;
                    worksheet.Cells[row, 8].Value = data.OPENING_QTY;
                    worksheet.Cells[row, 9].Value = data.OPENING_VAL;
                    worksheet.Cells[row, 10].Value = data.NET_PURCHASE_QTY;
                    worksheet.Cells[row, 11].Value = data.NET_PURCHASE_VAL;
                    worksheet.Cells[row, 12].Value = data.SALES_QTY;
                    worksheet.Cells[row, 13].Value = data.SALES_VAL;
                    worksheet.Cells[row, 14].Value = data.WASTAGE_QTY;
                    worksheet.Cells[row, 15].Value = data.WASTAGE_VAL;
                    worksheet.Cells[row, 16].Formula = "(H" + row.ToString() + "+J" + row.ToString() + ")-(L" + row.ToString() + "+N" + row.ToString() + ")";//=(H2+J2)-(L2+N2)
                    worksheet.Cells[row, 17].Formula = "P" + row.ToString() + "*G" + row.ToString();//=P2*G2
                    worksheet.Cells[row, 18].Value = data.CLOSING_QTY;
                    worksheet.Cells[row, 19].Formula = "R" + row.ToString() + "*G" + row.ToString();//=R2*G2
                    worksheet.Cells[row, 20].Formula = "H" + row.ToString() + "+J" + row.ToString() + "-L" + row.ToString() + "-N" + row.ToString() + "-R" + row.ToString();//=H2+J2-L2-N2-R2
                    worksheet.Cells[row, 21].Formula = "T" + row.ToString() + "*G" + row.ToString();//=T2*G2
                    worksheet.Cells[row, 22].Formula = "I" + row.ToString() + "+K" + row.ToString() + "-S" + row.ToString();//=I2+K2-S2
                    worksheet.Cells[row, 23].Formula = "M" + row.ToString() + "-V" + row.ToString();//=M2-V2
                    worksheet.Cells[row, 24].Formula = "IFERROR(W" + row.ToString() + "/M" + row.ToString() + "*100,0)";//=IFERROR(W2/M2*100,"0")
                    //worksheet.Cells["E"+row.ToString()].Formula ="A"+ row.ToString()+"&"+ "B" + row.ToString();
                    row++;
                }
                worksheet.Cells[row + 1, 7].Formula = "SUM(G2:G" + row.ToString() + ")";//=SUM(G2:G62)
                worksheet.Cells[row + 1, 8].Formula = "SUM(H2:H" + row.ToString() + ")";
                worksheet.Cells[row + 1, 9].Formula = "SUM(I2:I" + row.ToString() + ")";
                worksheet.Cells[row + 1, 10].Formula = "SUM(J2:J" + row.ToString() + ")";
                worksheet.Cells[row + 1, 11].Formula = "SUM(K2:K" + row.ToString() + ")";
                worksheet.Cells[row + 1, 12].Formula = "SUM(L2:L" + row.ToString() + ")";
                worksheet.Cells[row + 1, 13].Formula = "SUM(M2:M" + row.ToString() + ")";
                worksheet.Cells[row + 1, 14].Formula = "SUM(N2:N" + row.ToString() + ")";
                worksheet.Cells[row + 1, 15].Formula = "SUM(O2:O" + row.ToString() + ")";
                worksheet.Cells[row + 1, 16].Formula = "SUM(P2:P" + row.ToString() + ")";//=SUM(G2:G62)
                worksheet.Cells[row + 1, 17].Formula = "SUM(Q2:Q" + row.ToString() + ")";
                worksheet.Cells[row + 1, 18].Formula = "SUM(R2:R" + row.ToString() + ")";
                worksheet.Cells[row + 1, 19].Formula = "SUM(S2:S" + row.ToString() + ")";
                worksheet.Cells[row + 1, 20].Formula = "SUM(T2:T" + row.ToString() + ")";
                worksheet.Cells[row + 1, 21].Formula = "SUM(U2:U" + row.ToString() + ")";
                worksheet.Cells[row + 1, 22].Formula = "SUM(V2:V" + row.ToString() + ")";
                worksheet.Cells[row + 1, 23].Formula = "SUM(W2:W" + row.ToString() + ")";
                worksheet.Cells[row + 1, 24].Formula = "W" + (row+1).ToString() + "/M" + (row+1).ToString() + "*100";//=W222/M222*100
        }
    }
}
