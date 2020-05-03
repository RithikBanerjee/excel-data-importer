using System;
using System.Linq;
using ExcelDataImporter.Model;
using System.Collections.Generic;

namespace ExcelDataImporter.LightCellDataHandlers
{
    internal class DemoTableDataHandler : BaseLightCellDataHandler<List<DemoTable>>
    {
        private DemoTable RowData;
        private DemoTable TempData;
        public DemoTableDataHandler(Sheet<List<DemoTable>> sheet) : base(sheet)
        {
            //define or add construction of the DataType to needed Data
            sheet.ValidData = new List<DemoTable>();
            TempData = new DemoTable();
        }

        protected override bool StartEachRow()
        {
            //define or add construction of the DataType to row Data
            RowNumber = string.Empty;
            DataInvalidMessage = string.Empty;
            RowData = new DemoTable()
            {
                Details = new DemoDetails(),
                Items = new List<DemoItems>() { new DemoItems () }
            };
            return true;
        }

        protected override bool ProcessCellFurther(Column columnInfo, string cellValue)
        {
            //assign every cell to property of a row object
            object tableToInsert; Type tableType;
            switch (columnInfo.DBTableName)
            {
                case nameof(DemoTable):
                    tableType = typeof(DemoTable);
                    tableToInsert = RowData;
                    break;
                case nameof(DemoDetails):
                    tableType = typeof(DemoDetails);
                    tableToInsert = RowData.Details;
                    break;
                case nameof(DemoItems):
                    tableType = typeof(DemoItems);
                    tableToInsert = RowData.Items.First();
                    break;
                default:
                    return false;
            }
            var property = tableType.GetProperty(columnInfo.DBFieldName);
            try
            {
                property.SetValue(tableToInsert, Convert.ChangeType(cellValue, property.PropertyType));
            }
            catch (Exception)
            {
                property.SetValue(tableToInsert, Convert.ChangeType(columnInfo.DefaultValue, property.PropertyType));
            }
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

            //assign for 1-M or M-M realtion property
            if (Equals(TempData.Id, RowData.Id))
            {
                TempData.Items.Add(RowData.Items.First());
                if (Equals(Row.Index, Sheet.RowCount - 1))
                    Sheet.ValidData.Add(TempData);
                return true;
            }
            // then add the row data to valid data
            if (TempData.Id != 0)
                Sheet.ValidData.Add(TempData);

            TempData = RowData;
            return true;
        }
    }
}
