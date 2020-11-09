using System;
using Aspose.Cells;
using ExcelDataImporter.Model;
using System.Collections.Generic;
using ExcelDataImporter.Builder;
using ExcelDataImporter.LightCellDataHandlers;

namespace ExcelDataImporter.DataImporter
{
    //demo data importer for type: list of objects
    public class DemoTableImporter : BaseDataImporter<List<DemoTable>>
    {
        public DemoTableImporter(string excelFilePath, string excelSchemaPath) : base(excelFilePath, excelSchemaPath)
        {
        }

        public override bool ValidateData()
        {
            PreapareExcelBeforeDataValidation();
            SortRows("Id");

            var isDataValid = true;
            var builder = new DataBuilder();
            foreach(var sheet in Workbook.Sheets)
            {
                var handler = new DemoTableDataHandler(sheet);
                builder.GetDataFromExcel(sheet.Name, handler, Workbook.Path);
                if (sheet.InvalidData.Rows.Count != 0)
                    isDataValid = false;
            }
            return isDataValid;
        }

        private void SortRows( string sortColumn1)
        {
            var workbook = new Workbook(Workbook.Path);
            foreach(var sheet in Workbook.Sheets)
            {
                var worksheet = workbook.Worksheets[sheet.Name];
                var dataSorter = workbook.DataSorter;
                dataSorter.Order1 = SortOrder.Ascending;
                dataSorter.Key1 = sheet.Columns.Find(x => string.Equals(x.DBFieldName, sortColumn1, StringComparison.CurrentCultureIgnoreCase)).ColumnIndex;
                dataSorter.Sort(worksheet.Cells, 1, 0, worksheet.Cells.MaxRow, worksheet.Cells.MaxColumn);
            }
        }
    }
}
