using System;
using Aspose.Cells;
using ExcelDataImporter.Model;
using System.Text.RegularExpressions;

namespace ExcelDataImporter.LightCellDataHandlers
{
    //common class to handle indiviual excel cell
    public abstract class BaseLightCellDataHandler<T> : LightCellsDataHandler
    {
        protected Row Row;
        protected string RowNumber;
        protected readonly Sheet<T> Sheet;
        protected string DataInvalidMessage;
        protected BaseLightCellDataHandler(Sheet<T> sheet)
        {
            Sheet = sheet;
        }

        public bool StartSheet(Worksheet sheet)
        {
            sheet.Cells.DeleteBlankRows();
            sheet.Cells.DeleteBlankColumns();
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            if (rowIndex == 0)
                return false;

            return StartEachRow();
        }
        
        public bool ProcessRow(Row row)
        {
            Row = row;
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            var cellValue = Convert.ToString(cell.Value);

            //last cell of the row is rowIndex
            if(Equals(cell.Column, Sheet.Columns.Count))
            {
                RowNumber = cellValue;
                return ProcessRowFurther();
            }

            //finding column information as stored before
            var columnInfo = Sheet.Columns.Find(x => Equals(x.ColumnIndex, cell.Column));
            if (columnInfo == null)
                return false;

            //check and assign regex message if invalid cell data
            if (!Regex.IsMatch(cellValue, columnInfo.RegExPattern))
                DataInvalidMessage = $"{DataInvalidMessage} {columnInfo.RegexMessageIfInvalid}.";
            
            return ProcessCellFurther(columnInfo, cellValue);
        }

        protected abstract bool StartEachRow();

        protected abstract bool ProcessCellFurther(Model.Column columnInfo, string cellValue);

        protected abstract bool ProcessRowFurther();
        
    }
}
