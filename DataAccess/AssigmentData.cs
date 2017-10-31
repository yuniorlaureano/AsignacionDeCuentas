using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Entity;
using DataAccess.Repository;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using Entity.Results;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataAccess
{
    public class AssigmentData
    {
        OracleBasicOperation oracleOperation;

        public AssigmentData()
        {
            oracleOperation = new OracleBasicOperation();
        }

        public AssignmentResult InsertAssignment(List<Assignment> assignments, string userCode)
        {
            int count = 0;
            OracleParameter[] prm = 
            {
               new OracleParameter{ ParameterName = "in_sentencia", OracleDbType = OracleDbType.Int32},
               new OracleParameter{ ParameterName = "in_user_code", OracleDbType = OracleDbType.Int32, Value = userCode},
               new OracleParameter{ ParameterName = "out_fected_rows", OracleDbType = OracleDbType.Int16, Direction = System.Data.ParameterDirection.Output},
            };

            OracleTransaction transaction = oracleOperation.oracleConnection.BeginTransaction();

            try
            {
                oracleOperation.ExecuteNonQuery("sp_insert_assignment", prm, CommandType.StoredProcedure,DataAccess.Repository.Schema.YBRDS_PROD);
                count += Convert.ToSByte(prm[5].Value.ToString());
            }
            catch (OracleException except)
            {
                transaction.Rollback();
                oracleOperation.CloseConnection();
                return new AssignmentResult
                {
                    IsError = true,
                    StackTrace = except.StackTrace,
                    Message = except.Message
                };
            }

            if (count == (assignments.Count - 1))
            {
                transaction.Commit();
                oracleOperation.CloseConnection();
                return new AssignmentResult
                {
                    IsError = false,
                    Message = "Asignaciones inseratadas.",
                };
            }

            transaction.Rollback();
            oracleOperation.CloseConnection();
            return new AssignmentResult 
            {
                IsError = true,
                Message = "Error al insertar las asignaciones, presione intentar nuevamente.",
            };
        }
    }
}
