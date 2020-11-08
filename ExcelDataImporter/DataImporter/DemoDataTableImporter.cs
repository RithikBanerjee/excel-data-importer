using System.Data;
using ExcelDataImporter.Builder;
using ExcelDataImporter.LightCellDataHandlers;

namespace ExcelDataImporter.DataImporter
{
    //data validator class created based on the datatype
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
