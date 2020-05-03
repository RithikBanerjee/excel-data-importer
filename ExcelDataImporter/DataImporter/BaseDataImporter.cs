using System;
using System.IO;
using Aspose.Cells;
using ExcelDataImporter.Model;
using ExcelDataImporter.Context;

namespace ExcelDataImporter.DataImporter
{
    public abstract class BaseDataImporter<T>
    {
        public WorkbookSchema<T> Workbook;
        public BaseDataImporter(string excelFilePath, string excelSchemaPath)
        {
            Workbook = WorkBookSchemaContext.GetSchema<T>(excelSchemaPath);
            Workbook.Path = SaveExcel(excelFilePath);
        }

        private string SaveExcel(string excelFilePath)
        {
            var temporyPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\")) + "Excel";
            if (!Directory.Exists(temporyPath))
                Directory.CreateDirectory(temporyPath);

            temporyPath = temporyPath + @"\" + $"{DateTime.Now.Ticks}.xlsx";
            if (!File.Exists(temporyPath))
            {
                var myFile = File.Create(temporyPath);
                myFile.Close();
            }
            var workbook = new Workbook(excelFilePath);
            workbook.Save(temporyPath);
            return temporyPath;
        }

        public bool ValidateSchema()
        {
            var helper = new ExcelFileContext(Workbook.Path);
            var sheetPresents = helper.PrepareAndGetSheetsPresent(Workbook.Sheets);
            var context = new WorkBookSchemaContext();
            return context.ValidateSchema(sheetPresents, Workbook);
        }

        public abstract bool ValidateData();

        protected void PreapareExcelBeforeDataValidation()
        {
            var workbook = new Workbook(Workbook.Path);
            foreach(var sheet in Workbook.Sheets)
            {
                AssignRowIdBeforeSorting(workbook, sheet);
                AssignDateFormat(workbook, sheet);
            }
            workbook.Save(Workbook.Path);
        }

        private void AssignRowIdBeforeSorting(Workbook workbook, Sheet<T> sheet)
        {
            var worksheet = workbook.Worksheets[sheet.Name];
            worksheet.Cells.DeleteBlankRows();
            var lastColumnIndex = worksheet.Cells.MaxColumn + 1;
            for (int i = 1; i <= worksheet.Cells.MaxRow; i++)
            {
                var row = worksheet.Cells.Rows.GetRowByIndex(i);
                if (row != null)
                {
                    var cell = worksheet.Cells[row.Index, lastColumnIndex];
                    cell.PutValue(row.Index + 1);
                }
            }
        }

        private void AssignDateFormat(Workbook workbook, Sheet<T> sheet)
        {
            var worksheet = workbook.Worksheets[sheet.Name];
            var dateColumns = sheet.Columns.FindAll(x => Equals(x.DataType, typeof(DateTime)));
            foreach (var column in dateColumns)
                worksheet.Cells.Columns[column.ColumnIndex].ApplyStyle(
                    #pragma warning disable CS0618
                    style: new Style() { Number = 14, Custom = "mm-dd-yyyy" },
                    #pragma warning restore CS0618
                    flag: new StyleFlag { NumberFormat = true });
        }
    }
}
