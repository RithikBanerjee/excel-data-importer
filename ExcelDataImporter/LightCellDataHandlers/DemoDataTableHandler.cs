using System.Data;
using ExcelDataImporter.Model;

namespace ExcelDataImporter.LightCellDataHandlers
{
    internal class DemoDataTableHandler : BaseLightCellDataHandler<DataTable>
    {
        private DataRow RowData;
        public DemoDataTableHandler(Sheet<DataTable> sheet) : base(sheet)
        {
            //define or add construction of the DataType to needed Data
            sheet.ValidData = new DataTable(sheet.Name);
            sheet.Columns.ForEach(x => sheet.ValidData.Columns.Add(x.DBFieldName));
        }
        protected override bool StartEachRow()
        {
            //define or add construction of the DataType to row Data
            RowNumber = string.Empty;
            DataInvalidMessage = string.Empty;
            RowData = Sheet.ValidData.NewRow();
            return true;
        }
        protected override bool ProcessCellFurther(Column columnInfo, string cellValue)
        {
            //assign every cell to property of a row object
            RowData[columnInfo.DBFieldName] = cellValue;
            return true;
        }

        protected override bool ProcessRowFurther()
        {
            //if regex failed add row to invalid data
            if (!string.IsNullOrEmpty(DataInvalidMessage))
            {
                Sheet.InvalidData.Rows.Add(RowNumber, DataInvalidMessage);
                return true;
            }
            //add row-wise or parent-child validation here

            // then add the row data to valid data
            Sheet.ValidData.Rows.Add(RowData);
            return true;

        }
    }
}
