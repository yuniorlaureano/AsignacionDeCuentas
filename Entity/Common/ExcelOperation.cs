using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;

namespace Entity.Common
{
    public enum Provider
    {
        XLS,
        XLSX
    }

    public class ExcelOperation
    {
        public string[] Sheets { get; set; }
        public string FileExtension { get; set; }
        public string FileName { get; set; }
        public Provider Provider { get; set; }
        public OleDbConnection OldbCn { get; set; }
        public string ConnectionString { get; set; }
                
        /// <summary>
        /// Open the connection acording the the type of excel file. whether it is xls, or xlsx
        /// </summary>
        /// <param name="provider">Tipos of file, XLS or XLSX</param>
        /// <param name="fileName">The name of the excel file</param>
        public void OpenConnection(Provider provider, string fileName)
        {           
            switch(provider)
            {
                case Provider.XLS: this.ConnectionString = string.Format(ConfigurationManager.ConnectionStrings["XLS"].ToString(), "\"Excel 8.0;IMEX=1;HDR=Yes;TypeGuessRows=0;ImportMixedTypes=Text\"", fileName);
                    break;
                case Provider.XLSX: this.ConnectionString = string.Format(ConfigurationManager.ConnectionStrings["XLSX"].ToString(), "\"Excel 12.0;IMEX=1;HDR=Yes;TypeGuessRows=0;ImportMixedTypes=Text\"", fileName);
                    break;
                default: throw new  ArgumentException("Unown file extension: ["+provider.ToString()+"]");
            }

            try
            {
                if (this.OldbCn == null)
                {
                    this.OldbCn = new OleDbConnection(this.ConnectionString);
                    this.OldbCn.Open();
                }

                if (this.FileName != fileName || this.Provider != provider)
                {
                    this.CloseConnection();
                    this.OldbCn = new OleDbConnection(this.ConnectionString);
                    this.OldbCn.Open();
                }

                if (this.OldbCn.State == System.Data.ConnectionState.Closed)
                {
                    this.OldbCn.Open();
                }
            }
            catch (OleDbException except)
            {
                throw except;
            }
            finally
            {
                this.Provider = provider;
                this.FileName = fileName;
            }
        }

        /// <summary>
        /// Close the connection.
        /// </summary>
        public void CloseConnection()
        {
            try
            {
                if (this.OldbCn != null && this.OldbCn.State == System.Data.ConnectionState.Open)
                {
                    this.OldbCn.Close();
                }
            }
            catch (OleDbException except)
            {                
                throw except;
            }
        }

        /// <summary>
        /// Get the sheet names of the excel documents 
        /// </summary>
        /// <param name="provider">Tipos of file, XLS or XLSX</param>
        /// <param name="fileName">The name of the excel file</param>
        /// <returns>string[]</returns>
        public string[] GetSheetNames(Provider provider, string fileName)
        {
            try
            {
                this.OpenConnection(provider, fileName);
                DataTable sheets = this.OldbCn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                this.CloseConnection();
                int count = sheets.Rows.Count;
                this.Sheets = new string[count];

                for (int i = 0; i < count; i++)
                {
                    this.Sheets[i] = sheets.Rows[i]["TABLE_NAME"].ToString();
                }                
            }
            catch (OleDbException except)
            {
                throw except;
            }
            finally
            {
                this.CloseConnection();
            }

            return this.Sheets;
        }

        /// <summary>
        /// Execute a query that return a data reader, with the data contained in the excel.
        /// </summary>
        /// <param name="query">Sql senteces for being executed.</param>
        /// <param name="provider">Tipos of file, XLS or XLSX</param>
        /// <param name="fileName">The name of the excel file</param>
        /// <returns>OleDbDataReader</returns>
        public OleDbDataReader GetData(string query, Provider provider, string fileName)
        {
            OleDbDataReader reader;

            try
            {
                this.OpenConnection(provider, fileName);
                OleDbCommand command = new OleDbCommand(query, this.OldbCn);
                command.CommandType = System.Data.CommandType.Text;
                reader = command.ExecuteReader();
            }
            catch (OleDbException except)
            {
                throw except;
            }

            return reader;
        }
    
        /// <summary>
        /// Create the query for a select from a specified sheet name.
        /// </summary>
        /// <param name="query">Sql statement to be executed</param>
        /// <param name="sheet">The name fo the excel sheet</param>
        /// <returns>string</returns>
        public string Query(string query, string sheet)
        {
            return string.Format("{0} {1}", query, sheet);
        }
    
        /// <summary>
        /// set the provider, whether it is xls, or xlsx
        /// </summary>
        /// <param name="fileName">The name of the file been proccessed</param>
        public void SetProvider(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            
            switch (extension)
            {
                case ".xls": this.Provider = Provider.XLS;
                    break;
                case ".xlsx": this.Provider = Provider.XLSX;
                    break;
                default: throw new ArgumentException("La exgension ther archivo: ["+extension+"], no es valida.");
            }

            this.FileExtension = extension;
        }
    
    }
}
