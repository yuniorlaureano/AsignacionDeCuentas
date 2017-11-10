using DataAccess.Repository;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Entity;
using Entity.Results;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess
{
    public class SupervisorData
    {
        OracleBasicOperation oracleOperation;

        public SupervisorData()
        {
            oracleOperation = new OracleBasicOperation();
        }

        /// <summary>
        /// Permite obtner los supervisores
        /// </summary>
        /// <returns>List<Supervisor></returns>
        public List<Supervisor> GetSupervisors()
        {
            List<Supervisor> supervisors = new List<Supervisor>();
            OracleDataReader reader = null;

            OracleParameter[] prm = 
            { 
                new OracleParameter { ParameterName = "resultset", OracleDbType = OracleDbType.RefCursor, Direction = System.Data.ParameterDirection.Output}
            };

            try
            {
                reader = this.oracleOperation.ExecuteReader("pgk_account_assigment.get_supervisor", prm, CommandType.StoredProcedure, DataAccess.Repository.Schema.YBRDS_PROD);
                while (reader.Read())
                {
                    supervisors.Add(new Supervisor
                    {
                        Id = reader["supervisor_id"].ToString(),
                        Name = reader["supervisor"].ToString(),
                        EjecutiveCard = reader["tarjeta"].ToString(),
                        EjecutiveName = reader["ejecutivo"].ToString(),
                        Unit = reader["unidad"].ToString()
                    });
                }
            }
            catch (OracleException except)
            {
                throw except;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();    
                }

                this.oracleOperation.CloseConnection();
            }

            return supervisors;
        }

        public string WriteToExcel(List<Supervisor> supervisors, string sheetName, string fileName, string savingPath)
        {
            string archivo = savingPath + fileName + ".xlsx";

            using (SpreadsheetDocument workbook = SpreadsheetDocument.Create(archivo, SpreadsheetDocumentType.Workbook))
            {
                OpenXmlWriter writer;

                workbook.AddWorkbookPart();
                WorksheetPart wsp = workbook.WorkbookPart.AddNewPart<WorksheetPart>();

                writer = OpenXmlWriter.Create(wsp);
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                WriteExcelHeader(new string[] { "Supervisor Id", "Supervidor", "Tarjeta Ejecutivo", "Ejecutivo", "Unidad" }, writer);
                WriteExcelValues(supervisors, writer);

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

        public void WriteExcelValues(List<Supervisor> supervisors, OpenXmlWriter writer)
        {
            int currentRow = 0;

            List<OpenXmlAttribute> row;
            List<OpenXmlAttribute> cell = new List<OpenXmlAttribute> { new OpenXmlAttribute("t", null, "inlineStr") };

            for (int i = 1; i <= supervisors.Count; i++)
            {
                currentRow = i - 1;
                
                row = new List<OpenXmlAttribute> { new OpenXmlAttribute("r", null, (i + 1).ToString()) };
                writer.WriteStartElement(new Row(), row);

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(supervisors[currentRow].Id.ToString())));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(supervisors[currentRow].Name)));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(supervisors[currentRow].EjecutiveCard.ToString())));
                writer.WriteEndElement();

                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(supervisors[currentRow].EjecutiveName)));
                writer.WriteEndElement();
                
                writer.WriteStartElement(new Cell(), cell);
                writer.WriteElement(new InlineString(new Text(supervisors[currentRow].Unit.ToString())));
                writer.WriteEndElement();

                writer.WriteEndElement();
            }
        }

    }
}
