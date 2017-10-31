using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace DataAccess.Repository
{
    public class OracleBasicOperation
    {
        private ConnectionManager connectionManager;
        public OracleConnection oracleConnection;
        private Schema schema;

        public OracleBasicOperation()
        {
            connectionManager = new ConnectionManager();
        }

        public void OpenConnection(Schema schema)
        {
            try
            {
                if (this.oracleConnection == null)
                {
                    this.oracleConnection = new OracleConnection(connectionManager.SetConnectionString(schema));
                    this.schema = schema;
                }

                if (this.schema != schema)
                {
                    this.CloseConnection();
                    this.oracleConnection = new OracleConnection(connectionManager.SetConnectionString(schema));
                    this.schema = schema;
                }

                if (this.oracleConnection.State == System.Data.ConnectionState.Closed)
                {
                    this.oracleConnection.Open();
                }
            }
            catch (OracleException except)
            {
                throw except;
            }
        }

        public void CloseConnection()
        {
            try
            {
                if (this.oracleConnection != null && this.oracleConnection.State == System.Data.ConnectionState.Open)
                {
                    this.oracleConnection.Close();
                }
            }
            catch(OracleException except)
            {
                throw except;
            }
        }

        /// <summary>
        /// Ejecuta un comando, y retorna true.
        /// </summary>
        /// <param name="query">La sentencia sql a ejecutar</param>
        /// <param name="prm">Un arreglo de parametros</param>
        /// <param name="commandType">Si es procedure o directa</param>
        /// <param name="schema">A que base de datos se conectara</param>
        /// <returns>DataSet</returns>
        public DataSet GetDataSet(string query, OracleParameter[] prm, CommandType commandType, Schema schema)
        {
            DataSet resultset = new DataSet();

            try
            {               
                this.OpenConnection(schema);
                OracleCommand command = new OracleCommand(query, this.oracleConnection);
                command.CommandType = commandType;

                if (prm.Length > 0)
                {
                    command.Parameters.AddRange(prm);
                }

                OracleDataAdapter adapter = new OracleDataAdapter(command);
                adapter.Fill(resultset);               
            }
            catch (OracleException except)
            {
                throw except;
            }

            return resultset;
        }

        /// <summary>
        /// Ejecuta un comando, y retorna true.
        /// </summary>
        /// <param name="query">La sentencia sql a ejecutar</param>
        /// <param name="prm">Un arreglo de parametros</param>
        /// <param name="commandType">Si es procedure o directa</param>
        /// <param name="schema">A que base de datos se conectara</param>
        /// <returns>bool</returns>
        public bool ExecuteNonQuery(string query, OracleParameter[] prm, CommandType commandType, Schema schema)
        {
            bool resultset = false;

            try
            {
                this.OpenConnection(schema);
                OracleCommand command = new OracleCommand(query, this.oracleConnection);
                command.CommandType = commandType;

                if (prm.Length > 0)
                {
                    command.Parameters.AddRange(prm);
                }

                resultset = command.ExecuteNonQuery() > 0;
            }
            catch (OracleException except)
            {
                throw except;
            }

            return resultset;
        }

        /// <summary>
        /// Retorna un data reader.
        /// </summary>
        /// <param name="query">La sentencia sql a ejecutar</param>
        /// <param name="prm">Un arreglo de parametros</param>
        /// <param name="commandType">Si es procedure o directa</param>
        /// <param name="schema">A que base de datos se conectara</param>
        /// <returns>OracleDataReader</returns>
        public OracleDataReader ExecuteReader(string query, OracleParameter[] prm, CommandType commandType, Schema schema)
        {
            OracleDataReader resultset = null;

            try
            {
                this.OpenConnection(schema);
                OracleCommand command = new OracleCommand(query, this.oracleConnection);
                command.CommandType = commandType;

                if (prm.Length > 0)
                {
                    command.Parameters.AddRange(prm);
                }

                resultset = command.ExecuteReader();
            }
            catch (OracleException except)
            {
                throw except;
            }

            return resultset;
        }        
    }
}
