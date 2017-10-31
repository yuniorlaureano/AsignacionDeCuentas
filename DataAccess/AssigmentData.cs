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
using Entity.Common;
using System.Data.OleDb;

namespace DataAccess
{
    public class AssigmentData
    {
        OracleBasicOperation oracleOperation;
        ExcelOperation excelOp;

        public AssigmentData()
        {
            oracleOperation = new OracleBasicOperation();
            excelOp = new ExcelOperation();
        }

        public AssignmentResult InsertAssignment(StatementResult statement, string userCode)
        {
            int count = 0;
            OracleParameter[] prm = 
            {
               new OracleParameter{ ParameterName = "in_sentencia", OracleDbType = OracleDbType.Varchar2, Value = ""},
               new OracleParameter{ ParameterName = "in_user_code", OracleDbType = OracleDbType.Int32, Value = userCode},
               new OracleParameter{ ParameterName = "out_fected_rows", OracleDbType = OracleDbType.Int16, Direction = System.Data.ParameterDirection.Output},
            };

            OracleTransaction transaction = oracleOperation.oracleConnection.BeginTransaction();

            try
            {
                foreach (string st in statement.Statements)
                {
                    prm[0].Value = st;
                    oracleOperation.ExecuteNonQuery("pgk_account_assigment.sp_insert_assignments", prm, CommandType.StoredProcedure, DataAccess.Repository.Schema.YBRDS_PROD);
                    count += Convert.ToInt32(prm[2].Value.ToString());
                }
            }
            catch (OracleException except)
            {
                transaction.Rollback();
                oracleOperation.CloseConnection();
                return new AssignmentResult
                {
                    IsError = true,
                    StackTrace = except.StackTrace,
                    Message = except.Message,
                    InsertedRows = count,
                    SpectedRows = statement.RowCount
                };
            }
            finally
            {
                oracleOperation.CloseConnection();
            }

            if (count == statement.RowCount)
            {
                transaction.Commit();
                oracleOperation.CloseConnection();
                return new AssignmentResult
                {
                    IsError = false,
                    Message = "[" + statement.RowCount + "] Asignaciones inseratadas.",
                    InsertedRows = statement.RowCount,
                    SpectedRows = statement.RowCount
                };
            }

            transaction.Rollback();
            oracleOperation.CloseConnection();
            return new AssignmentResult 
            {
                IsError = true,
                Message = "Error al insertar las asignaciones, presione intentar nuevamente.",
                InsertedRows = count,
                SpectedRows = statement.RowCount
            };
        }

        public StatementResult GenerateAssigncSentences(string fileNameWithLocation, Provider provider, string sheet)
        {
            StatementResult statement = new StatementResult();
            statement.Statements = new List<string>();
            string sentencia = "";
            int count = 0;
         
            OleDbDataReader rs = null;
            string query = excelOp.Query("SELECT SUBSCR_ID, CANV_CODE, CANV_EDITION, EMPLOYEE_ID FROM", sheet);

            try
            {
                rs = excelOp.GetData(query, provider, fileNameWithLocation);
                while(rs.Read())
                {
                    sentencia += string.Format("SELECT '{0}' SUBSCR_ID, '{1}' CANV_CODE, '{2}' CANV_EDITION, '{3}' EMPLOYEE_ID FROM DUAL ", rs["SUBSCR_ID"], rs["CANV_CODE"], rs["CANV_EDITION"], rs["EMPLOYEE_ID"]);

                    count += 1;

                    if (count == 20)
                    {
                        statement.Statements.Add(sentencia.Substring(0, sentencia.Length - 7));
                        sentencia = string.Empty;
                        count = 0;                        
                    }

                    statement.RowCount += 1;
                }
            }
            catch (OleDbException except)
            {                
                throw except;
            }
            finally
            {
                rs.Close();
                excelOp.CloseConnection();
            }

            return statement;
        }
    }
}
