using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;
//using Microsoft.Data.Schema.ScriptDom;
//using Microsoft.Data.Schema.ScriptDom.Sql;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.IO;

namespace QueryGen
{
    public class QueryProcessor
    {
        private static string comma = ",";
        private static string singleQuote = "\'";

        public static string GenerateQuery(QueryBuilderInput qbInput)
        {
            StringBuilder queryInsertPart = new StringBuilder();
            StringBuilder queryValuesPart = new StringBuilder();
            StringBuilder finalQuery = new StringBuilder();
            string queryFirstHalf = string.Empty;
            string querySecondHalf = string.Empty;

            var wb = new XLWorkbook(qbInput.File);

            IXLWorksheet ws1;
            bool canGetWS1 = wb.Worksheets.TryGetWorksheet(qbInput.SheetName, out ws1);

            if (canGetWS1)
            {
                //Insert part
                queryInsertPart.Append(" INSERT INTO " + qbInput.TableNameDB.Trim() + "(");
                foreach (Column col in qbInput.ExcelColumns)
                {
                    queryInsertPart.Append(col.ColumnNameDB + comma);
                }

                queryFirstHalf = queryInsertPart.ToString().TrimEnd(',') + ") VALUES(";

                string value = "";
                //Values part
                for (int i = qbInput.StartRow; i <= qbInput.EndRow; i++)
                {
                    foreach (Column col in qbInput.ExcelColumns)
                    {
                        value = "";
                        if (ws1.Cell(col.ColumnNameExcel + i.ToString()).Value.ToString() == "" && col.PassNull)
                        {
                            value = "NULL";
                            queryValuesPart.Append(value + comma);
                        }
                        else
                        {
                            value = ws1.Cell(col.ColumnNameExcel + i.ToString()).Value.ToString();

                            if (col.IsString)

                                queryValuesPart.Append(singleQuote + value + singleQuote + comma);
                            else
                                queryValuesPart.Append(value + comma);
                        }


                    }

                    finalQuery.Append(queryFirstHalf + queryValuesPart.ToString().TrimEnd(',') + ") \r\n");
                    queryValuesPart = new StringBuilder();

                }
            }

            return finalQuery.ToString();
        }

        public static bool ValidateQuery(string sql, out string message)
        {
            message = string.Empty;
            bool hasErrors = false;
            StringBuilder sb = new StringBuilder();

            //string sql = "SELECT * FROM SomeTable WHERE (1=1";
            var p = new TSql100Parser(true);
            IList<ParseError> errors;

            p.Parse(new StringReader(sql), out errors);


            if (errors.Count == 0)
                Console.Write("No Errors");
            else
            {
                hasErrors = true;
                foreach (ParseError parseError in errors)
                    sb.Append(parseError.Message);

                message = sb.ToString();
            }

            return hasErrors;
        }
    }
}
