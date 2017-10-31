using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace DataAccess.Repository
{
    public enum Schema
    {
        SFA,
        BRDS_PROD,
        YBRDS_PROD,
        GESTION
    }

    public class ConnectionManager
    {
        public string OracleConnection { get; set; }

        /// <summary>
        ///  return the connection concerning to oracle.
        /// </summary>
        /// <param name="schema"></param>
        /// <returns>string</returns>
        public string SetConnectionString(Schema schema)
        {

            switch (schema)
            {

                case Schema.SFA:
                    {
                        OracleConnection = ConfigurationManager.ConnectionStrings["SFA_ConnectionString"].ToString();
                        break;
                    }
                case Schema.BRDS_PROD:
                    {
                        OracleConnection = ConfigurationManager.ConnectionStrings["BRDS_PROD_ConnectionString"].ToString();
                        break;
                    }
                case Schema.YBRDS_PROD:
                    {
                        OracleConnection = ConfigurationManager.ConnectionStrings["YBRDS_PROD_ConnectionString"].ToString();
                        break;
                    }
                case Schema.GESTION:
                    {
                        OracleConnection = ConfigurationManager.ConnectionStrings["GESTION_ConnectionString"].ToString();
                        break;
                    }
            }

            return OracleConnection;
        }

    }
}
