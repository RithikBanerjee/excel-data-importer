using System;
using System.Data;
using System.Collections.Generic;

namespace ExcelDataImporter.Model
{
    //excel schema resposne object
    public class WorkbookSchema<T>
    {
        public string Path { get; set; }
        public List<Sheet<T>> Sheets { get; set; }
        public DataTable InvalidSchema { get; set; }
    }

    public class Sheet<T>
    {
        public string Name { get; set; }
        public bool Valid { get; set; } = true;
        public List<Column> Columns { get; set; }
        public DataTable InvalidData { get; set; }
        public T ValidData { get; set; }
        public int RowCount { get; set; }
        public List<DBTables> LinkedDBTables { get; set; }
    }

    public class Column
    {
        public List<string> NamesToFind { get; set; }
        public string ColumnName { get; set; }
        public int ColumnIndex { get; set; }
        public int RowIndex { get; set; }
        public string DBFieldName { get; set; }
        public string DBTableName { get; set; }
        public string ParentTable { get; set; }
        public bool Required { get; set; }
        public Type DataType { get; set; }
        public string DefaultValue { get; set; }
        public bool Found { get; set; }
        public string RegExPattern { get; set; }
        public string RegexMessageIfInvalid { get; set; }
    }

    public class DBTables
    {
        public string TableName { get; set; }
    }
}
