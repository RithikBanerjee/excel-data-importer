using System;
using System.Linq;
using System.Data;
using System.Data.CData.Excel;
using System.Collections.Generic;
using ExcelDataImporter.Model;

namespace ExcelDataImporter.Context
{
    internal class ExcelFileContext
    {
        private readonly string ConnectionString;

        internal ExcelFileContext(string excelFilePath)
        {
            ConnectionString = $"Excel File={excelFilePath};";
        }

        private ExcelConnection GetConnection(bool header)
        {
            return header ? new ExcelConnection(ConnectionString) : new ExcelConnection($"AutoCache=false;{ConnectionString}header=false;");
        }

        internal List<string> PrepareAndGetSheetsPresent<T>(ICollection<Sheet<T>> sheets)
        {
            var connection = GetConnection(false);
            connection.Open();
            var sheetsPresent = new List<string>();
            var excelSheets = GetSheetsInExcelFile(connection);
            foreach (var sheet in sheets)
            {
                var excelSheet = GetSheetNameIfFound(excelSheets, sheet.Name);
                if (excelSheet == null)
                    continue;

                sheet.Name = excelSheet;
                var headerRowId = AssignRownAndColumnIndex(connection, sheet.Name, sheet.Columns);
                if (headerRowId == 0)
                    continue;

                sheetsPresent.Add(sheet.Name);
                DeleteBlankRowsTillHeader(connection, sheet.Name, headerRowId);
                sheet.RowCount = GetRowCount(connection, sheet.Name);
            }
            connection.Close();
            return sheetsPresent;
        }

        private static int GetRowCount(ExcelConnection connection, string sheetName)
        {
            var data = new DataTable();
            var dataAdapter = new ExcelDataAdapter($"SELECT count(*) from [{sheetName}]", connection);
            dataAdapter.Fill(data);
            return Convert.ToInt32(data.Rows[0][0]);
        }

        private static List<string> GetSheetsInExcelFile(ExcelConnection connection)
        {
            var sheetsPresent = new List<string>();
            var excelSheets = connection.GetSchema("Tables");
            foreach (DataRow row in excelSheets.Rows)
            {
                sheetsPresent.Add(row["TABLE_NAME"].ToString());
            }
            return sheetsPresent;
        }

        private static string GetSheetNameIfFound(IEnumerable<string> excelSheets, string sheetTofind)
        {
            return excelSheets.FirstOrDefault(excelSheet =>
                string.Equals(excelSheet.Trim(), sheetTofind.Trim(), StringComparison.CurrentCultureIgnoreCase));
        }

        private static int AssignRownAndColumnIndex(ExcelConnection connection, string sheetName, List<Column> columns)
        {
            var data = new DataTable();
            var dataAdapter = new ExcelDataAdapter($"SELECT top 100 * from [{sheetName}]", connection);
            dataAdapter.Fill(data);

            foreach (DataRow dataRow in data.Rows)
            {
                var headerRowId = UpdateRowAndColumIndex(dataRow, columns);
                if (headerRowId != 0)
                    return headerRowId;
            }

            return 0;
        }

        private static int UpdateRowAndColumIndex(DataRow dataRow, List<Column> columns)
        {
            var rowIndex = (int)dataRow["RowId"];
            var rowItems = dataRow.ItemArray.ToList();
            foreach (var column in columns)
            {
                foreach (var columnName in column.NamesToFind)
                {
                    for (var i = 0; i < rowItems.Count; i++)
                    {
                        var item = rowItems[i].ToString().Trim();
                        if (item != columnName) continue;
                        column.ColumnName = columnName;
                        column.ColumnIndex = i - 1;
                        column.RowIndex = rowIndex;
                        column.Found = true;
                        break;
                    }
                    if (column.Found) break;
                }
            }

            return rowIndex;
        }

        private static void DeleteBlankRowsTillHeader(ExcelConnection connection, string sheetName, int headerRowId)
        {
            var deleteCommand = new ExcelCommand($"delete from [{sheetName}] where RowId='1' ", connection);
            for (var i = 1; i < headerRowId; i++)
                deleteCommand.ExecuteNonQuery();
        }
    }
}