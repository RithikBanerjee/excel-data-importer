using System;
using System.IO;
using System.Data;
using System.Linq;
using Newtonsoft.Json;
using ExcelDataImporter.Model;
using System.Collections.Generic;

namespace ExcelDataImporter.Context
{
    internal class WorkBookSchemaContext
    {
        internal static WorkbookSchema<T> GetSchema<T>(string schemaPath)
        {
            var jsonSchema = File.ReadAllText(schemaPath);
            return JsonConvert.DeserializeObject<WorkbookSchema<T>>(jsonSchema);
        }
        private void RemoveSheetsNotFound<T>(List<string> sheetsPresentInExcelFile, ICollection<Sheet<T>> sheetsRquired)
        {
            var sheetsNotFound = (from sheet in sheetsRquired
                                  let sheetFound = sheetsPresentInExcelFile.AsEnumerable()
                                      .Any(row => string.Equals(sheet.Name.Trim(), row.ToString().Trim(),StringComparison.CurrentCultureIgnoreCase))
                                  where !sheetFound
                                  select sheet);
            RemoveBlankOrInvalidSheets(sheetsRquired, sheetsNotFound);
        }

        private void MarkIfSheetIsValid<T>(IEnumerable<Sheet<T>> sheets)
        {
            foreach (var sheet in sheets)
            {
                foreach (var column in sheet.Columns)
                {
                    if (column.Required && !column.Found)
                    {
                        sheet.Valid = false;
                        break;
                    }
                }
                if (sheet.Valid)
                    RemoveOptionalColumnsNotFound(sheet);
            }
        }

        internal bool ValidateSchema<T>(List<string> sheetsPresent, WorkbookSchema<T> workbook)
        {
            workbook.InvalidSchema = new DataTable();
            workbook.InvalidSchema.Columns.Add("Error", typeof(string));
            workbook.InvalidSchema.Columns.Add("WhyInvalid", typeof(string));
            RemoveSheetsNotFound(sheetsPresent, workbook.Sheets);
            MarkIfSheetIsValid(workbook.Sheets);
            if (!workbook.Sheets.Any())
            {
                workbook.InvalidSchema.Rows.Add("N/A", "No Sheet found");
                return false;
            }
            var IsSchemaValid = true;
            foreach (var sheet in workbook.Sheets)
            {
                if (sheet.Valid)
                    continue;

                var columnNotFound = sheet.Columns.Where(c => !c.Found && c.Required).ToList();
                workbook.InvalidSchema.Rows.Add(sheet.Name, $"following required columns were not found: '{string.Join("','", columnNotFound.Select(c => c.NamesToFind.First()).ToList())}'");
                IsSchemaValid = false;
            }
            CreateAndAssignDataTable(workbook);
            return IsSchemaValid;
        }
        

        private static void RemoveOptionalColumnsNotFound<T>(Sheet<T> schemaSheet)
        {
            schemaSheet.Columns.RemoveAll(c => !c.Found && !c.Required);
        }

        private static void RemoveBlankOrInvalidSheets<T>(ICollection<Sheet<T>> sheets,
            IEnumerable<Sheet<T>> sheetsToRemove)
        {
            foreach (var sheet in sheetsToRemove)
                sheets.Remove(sheet);
        }

        private void CreateAndAssignDataTable<T>(WorkbookSchema<T> workbook)
        {
            foreach (var sheet in workbook.Sheets)
            {
                sheet.InvalidData = new DataTable(sheet.Name);
                sheet.InvalidData.Columns.Add("RowNumber");
                sheet.InvalidData.Columns.Add("WhyInvalid");
                
                /*
                sheet.InvalidData = new DataTable(sheet.Name);
                sheet.InvalidData.Columns.Add("RowNumber");
                var lastIndex = 0;
                foreach (var column in sheet.Columns)
                {
                    sheet.InvalidData.Columns.Add(column.DBFieldName);
                    lastIndex = column.ColumnIndex;
                }

                sheet.Columns.Add(new Column
                {
                    ColumnIndex = lastIndex + 1,
                    DBFieldName = "RowNumber",
                    IsValueOptional = false
                });
                sheet.InvalidData.Columns.Add("WhyInvalid");
                */
            }
        }
    }
}