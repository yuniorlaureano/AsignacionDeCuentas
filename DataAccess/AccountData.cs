using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataAccess.Repository;
using Entity;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Entity.Results;

namespace DataAccess
{
    public class AccountData
    {
        OracleBasicOperation oracleOperation;

        public AccountData()
        {
            oracleOperation = new OracleBasicOperation();
        }

        public List<Account> GetInvalid(string UserCode)
        {
            List<Account> invalidAccounts = null;
            OracleDataReader resultset = null;

            OracleParameter[] prm = 
            { 
                new OracleParameter { ParameterName = "in_user_code", OracleDbType = OracleDbType.Int32, Value = UserCode},
                new OracleParameter { ParameterName = "resultset", OracleDbType = OracleDbType.RefCursor, Direction = System.Data.ParameterDirection.Output}
            };

            resultset = this.oracleOperation.ExecuteReader("pgk_account_assigment.sp_get_invalid_account", prm, CommandType.StoredProcedure, DataAccess.Repository.Schema.YBRDS_PROD);

            while (resultset.Read())
            {
                invalidAccounts.Add(
                    new Account
                    {
                        CanvCode = resultset["CANV_CODE"].ToString(),
                        ChargeIn = Convert.ToInt32(resultset["CHARGE_IN"].ToString()),
                        ChargeOut = Convert.ToInt32(resultset["CHARGE_OUT"].ToString()),
                        Status = resultset["ACCOUNT_STATUS"].ToString(),
                        SubscrId = Convert.ToInt32(resultset["SUBSCR_ID"].ToString()),
                        CanvEdition = Convert.ToInt32(resultset["canv_edition"].ToString()),
                        UserAssigned = resultset["assigned_employee"].ToString(),
                        UserToAssign = resultset["to_assign_employee"].ToString()
                    });
            }

            return invalidAccounts;
        }

        public CountResult ValidInvalidAccount(string UserCode)
        {            
            OracleParameter[] prm = 
            { 
                new OracleParameter { ParameterName = "in_user_code", OracleDbType = OracleDbType.Int32, Value = UserCode},
                new OracleParameter { ParameterName = "out_invalid_count", OracleDbType = OracleDbType.Int32, Direction = System.Data.ParameterDirection.Output},
                new OracleParameter { ParameterName = "out_valid_count", OracleDbType = OracleDbType.Int32, Direction = System.Data.ParameterDirection.Output}
            };

            oracleOperation.ExecuteNonQuery("pgk_account_assigment.sp_get_account_count", prm, CommandType.StoredProcedure, Repository.Schema.YBRDS_PROD);
            oracleOperation.CloseConnection();

            return new CountResult {
                Valid = Convert.ToInt32(prm[1].Value), 
                Invalid = Convert.ToInt32(prm[2].Value)
            };
        }

        public string WriteToExcel(List<Account> assignments, string sheetName, string fileName, string savingPath)
        {
            string archivo = savingPath + fileName + ".xlsx";

            using (SpreadsheetDocument workbook = SpreadsheetDocument.Create(archivo, SpreadsheetDocumentType.Workbook))
            {
                OpenXmlWriter writer;

                workbook.AddWorkbookPart();
                WorksheetPart wsp = workbook.AddNewPart<WorksheetPart>();

                writer = OpenXmlWriter.Create(wsp);
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                WriteExcelHeader(new string[] { "Subscr Id", "Canv Code", "Canv Edition", "Charge In", "Charge Out", "Status", "Usuario Asignado", "Usuario a Asignar" }, writer);

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.Close();

                writer = OpenXmlWriter.Create(workbook.WorkbookPart);
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                writer.WriteElement(new Sheet()
                {
                    Name = sheetName,
                    SheetId = 1,
                    Id = workbook.WorkbookPart.GetIdOfPart(wsp)
                });

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.Close();

                workbook.Close();
            }

            return archivo;
        }

        public void WriteExcelHeader(string[] header, OpenXmlWriter writer)
        {
            List<OpenXmlAttribute> row = new List<OpenXmlAttribute> { new OpenXmlAttribute("r", null, "1") };
            writer.WriteStartElement(new Row(), row);

            List<OpenXmlAttribute> cell = new List<OpenXmlAttribute> { new OpenXmlAttribute("t", null, "inlineStr") };

            foreach (string h in header)
            {
                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(h)));
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public void WriteExcelValues(List<Account> assignments, OpenXmlWriter writer)
        {
            int currentRow = 0;

            List<OpenXmlAttribute> row = new List<OpenXmlAttribute>();
            List<OpenXmlAttribute> cell = new List<OpenXmlAttribute> { new OpenXmlAttribute("t", null, "inlineStr") };

            for (int i = 1; i <= assignments.Count; i++)
            {
                currentRow = i - 1;
                row[0] = new OpenXmlAttribute("r", null, (i + 1).ToString());
                writer.WriteStartElement(new Row(), row);

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].SubscrId.ToString())));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].CanvCode)));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].CanvEdition.ToString())));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].ChargeIn.ToString())));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].ChargeOut.ToString())));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].Status)));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].UserAssigned)));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(assignments[currentRow].UserToAssign)));
                writer.WriteEndElement();

                writer.WriteEndElement();
            }
        }
        
    }
}
