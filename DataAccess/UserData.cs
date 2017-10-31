using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Entity;
using DataAccess.Repository;
using System.Data.SqlClient;
using System.Data;

namespace DataAccess
{
    public class UserData
    {
        private SqlBasicOperation sqlBasicOperation;

        public UserData()
        {
            sqlBasicOperation = new SqlBasicOperation();
        }

        public User GetUserCode(string userName)
        {
            User user = null;
            SqlDataReader resultset = null;

            SqlParameter[] prm = 
            { 
                new SqlParameter { ParameterName = "@user_name", SqlDbType = SqlDbType.VarChar, Value = userName}
            };

            resultset = this.sqlBasicOperation.ExecuteReader("ASC_APP.GET_USER_CREDENTIAL", prm, CommandType.StoredProcedure, DataAccess.Repository.Schema.GESTION);

            while (resultset.Read())
            {
                user = new User
                {
                    UserCode = Convert.ToInt32(resultset["usr_codigo"].ToString())
                };
            }

            sqlBasicOperation.CloseConnection();

            return user;
        }
    }
}
