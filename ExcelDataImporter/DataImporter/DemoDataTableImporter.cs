using System.Data;
using ExcelDataImporter.Builder;
using ExcelDataImporter.LightCellDataHandlers;

namespace ExcelDataImporter.DataImporter
{
    //demo data importer for type: datatable
    public class DemoDataTableImporter : BaseDataImporter<DataTable>
    {
        public DemoDataTableImporter(string excelFilePath, string excelSchemaPath) : base(excelFilePath, excelSchemaPath)
        {
        }

        public override bool ValidateData()
        {
            PreapareExcelBeforeDataValidation();

            var isDataValid = true;
            var builder = new DataBuilder();
            foreach (var sheet in Workbook.Sheets)
            {
                var handler = new DemoDataTableHandler(sheet);
                builder.GetDataFromExcel(sheet.Name, handler, Workbook.Path);
                if (sheet.InvalidData.Rows.Count != 0)
                    isDataValid = false;
            }
            return isDataValid;
        }
    }
}
