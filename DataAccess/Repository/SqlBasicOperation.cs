using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace DataAccess.Repository
{
    public class SqlBasicOperation
    {
        private ConnectionManager connectionManager;
        public SqlConnection SqlCon;
        private Schema schema;

        public SqlBasicOperation()
        {
            connectionManager = new ConnectionManager();
        }

        public void OpenConnection(Schema schema)
        {
            try
            {
                if (this.SqlCon == null)
                {
                    this.SqlCon = new SqlConnection(connectionManager.SetConnectionString(schema));
                    this.schema = schema;
                }

                if (this.schema != schema)
                {
                    this.CloseConnection();
                    this.SqlCon = new SqlConnection(connectionManager.SetConnectionString(schema));
                    this.schema = schema;
                }

                if (this.SqlCon.State == System.Data.ConnectionState.Closed)
                {
                    this.SqlCon.Open();
                }
            }
            catch (SqlException except)
            {
                throw except;
            }
        }

        public void CloseConnection()
        {
            try
            {
                if (this.SqlCon != null && this.SqlCon.State == System.Data.ConnectionState.Open)
                {
                    this.SqlCon.Close();
                }
            }
            catch(SqlException except)
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
        public DataSet GetDataSet(string query, SqlParameter[] prm, CommandType commandType, Schema schema)
        {
            DataSet resultset = new DataSet();

            try
            {               
                this.OpenConnection(schema);
                SqlCommand command = new SqlCommand(query, this.SqlCon);
                command.CommandType = commandType;

                if (prm != null)
                {
                    command.Parameters.AddRange(prm);
                }

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(resultset);               
            }
            catch (SqlException except)
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
        public bool ExecuteNonQuery(string query, SqlParameter[] prm, CommandType commandType, Schema schema)
        {
            bool resultset = false;

            try
            {
                this.OpenConnection(schema);
                SqlCommand command = new SqlCommand(query, this.SqlCon);
                command.CommandType = commandType;

                if (prm != null)
                {
                    command.Parameters.AddRange(prm);
                }

                resultset = command.ExecuteNonQuery() > 0;
            }
            catch (SqlException except)
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
        /// <returns>SqlDataReader</returns>
        public SqlDataReader ExecuteReader(string query, SqlParameter[] prm, CommandType commandType, Schema schema)
        {
            SqlDataReader resultset = null;

            try
            {
                this.OpenConnection(schema);
                SqlCommand command = new SqlCommand(query, this.SqlCon);
                command.CommandType = commandType;

                if (prm != null)
                {
                    command.Parameters.AddRange(prm);
                }

                resultset = command.ExecuteReader();
            }
            catch (SqlException except)
            {
                throw except;
            }

            return resultset;
        }        
    }
}
