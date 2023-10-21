using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WareHouseProdcutApi.Models;

namespace WareHouseProdcutApi.Services
{
    public class WhPrdExpDateService
    {
        private string connectionString;
        public WhPrdExpDateService(string connectionString)
        {
            this.connectionString = connectionString;
        }
        public List<WhProduct> GetWhProducts(long product)
        {
            List<WhProduct> WhProducts = new List<WhProduct>();

            using(OracleConnection connection=new OracleConnection(connectionString))
            {
                connection.Open();
                string query = @"SELECT A.*,
                                        (SELECT SU_DESCRIPTION FROM GRAND_PRD_MASTER_FULL_NEW WHERE BARCODE=OB_BARCODE)SU_DESCRIPTION
                                FROM
                                    OWN_BRAND_EXPIRY_DATE A 
                                WHERE 
                                     OB_BARCODE LIKE '%"+ product + @"%' OR 
                                     OB_PROD LIKE '%"+product+"%'";
                using(OracleCommand cmd=new OracleCommand(query, connection))
                {
                    using(OracleDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                WhProduct whProduct = new WhProduct();
                                whProduct.ArtCode = int.Parse(reader["OB_PROD"].ToString());
                                whProduct.Barcode = long.Parse(reader["OB_BARCODE"].ToString());
                                whProduct.ExpDate = DateTime.Parse(reader["OB_EXP_DATE"].ToString());
                                whProduct.Su = int.Parse(reader["OB_SU"].ToString());
                                whProduct.SuDescription = reader["SU_DESCRIPTION"].ToString();
                                WhProducts.Add(whProduct);
                            }
                        }
                    }
                }
                connection.Close();
            }
            
            return WhProducts;
        }
    }
}
