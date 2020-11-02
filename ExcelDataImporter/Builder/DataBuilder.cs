using System;
using System.IO;
using Aspose.Cells;
using ExcelDataImporter.LightCellDataHandlers;
using static Aspose.Cells.LoadDataFilterOptions;

namespace ExcelDataImporter.Builder
{
    public class DataBuilder
    {
        //C Data license
        private const string LData = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz4NCjxMaWNlbnNlPg0KICAgIDxEYXRhPg0KICAgICAgICA8TGljZW5zZWRUbz5pckRldmVsb3BlcnMuY29tPC9MaWNlbnNlZFRvPg0KICAgICAgICA8RW1haWxUbz5pbmZvQGlyRGV2ZWxvcGVycy5jb208L0VtYWlsVG8+DQogICAgICAgIDxMaWNlbnNlVHlwZT5EZXZlbG9wZXIgT0VNPC9MaWNlbnNlVHlwZT4NCiAgICAgICAgPExpY2Vuc2VOb3RlPkxpbWl0ZWQgdG8gMTAwMCBkZXZlbG9wZXIsIHVubGltaXRlZCBwaHlzaWNhbCBsb2NhdGlvbnM8L0xpY2Vuc2VOb3RlPg0KICAgICAgICA8T3JkZXJJRD43ODQzMzY0Nzc4NTwvT3JkZXJJRD4NCiAgICAgICAgPFVzZXJJRD4xMTk0NDkyNDM3OTwvVXNlcklEPg0KICAgICAgICA8T0VNPlRoaXMgaXMgYSByZWRpc3RyaWJ1dGFibGUgbGljZW5zZTwvT0VNPg0KICAgICAgICA8UHJvZHVjdHM+DQogICAgICAgICAgICA8UHJvZHVjdD5Bc3Bvc2UuVG90YWwgUHJvZHVjdCBGYW1pbHk8L1Byb2R1Y3Q+DQogICAgICAgIDwvUHJvZHVjdHM+DQogICAgICAgIDxFZGl0aW9uVHlwZT5FbnRlcnByaXNlPC9FZGl0aW9uVHlwZT4NCiAgICAgICAgPFNlcmlhbE51bWJlcj57RjJCOTcwNDUtMUIyOS00QjNGLUJENTMtNjAxRUZGQTE1QUE5fTwvU2VyaWFsTnVtYmVyPg0KICAgICAgICA8U3Vic2NyaXB0aW9uRXhwaXJ5PjIwOTkxMjMxPC9TdWJzY3JpcHRpb25FeHBpcnk+DQogICAgICAgIDxMaWNlbnNlVmVyc2lvbj4zLjA8L0xpY2Vuc2VWZXJzaW9uPg0KICAgIDwvRGF0YT4NCiAgICA8U2lnbmF0dXJlPlFYTndiM05sTGxSdmRHRnNMb1B5YjJSMVkzUWdSbUZ0YVd4NTwvU2lnbmF0dXJlPg0KPC9MaWNlbnNlPg==";

        public DataBuilder()
        {
            var stream = new MemoryStream(Convert.FromBase64String(LData));
            stream.Seek(0, SeekOrigin.Begin);
            var license = new License();
            license.SetLicense(stream);
        }
        //data importing method
        public void GetDataFromExcel<T>(string sheetName, BaseLightCellDataHandler<T> dataHandler, string excelFilePath)
        {
            var options = new LoadOptions(LoadFormat.Xlsx)
            {
                LoadFilter = new CustomLoad(sheetName),
                LightCellsDataHandler = dataHandler
            };

            var workbook = new Workbook(excelFilePath, options);
            workbook.Dispose();
        }
    }
    //data loading settings class
    public class CustomLoad : LoadFilter
    {
        private readonly string SheetName;

        public CustomLoad(string sheetName)
        {

            SheetName = sheetName;
        }

        public override void StartSheet(Worksheet sheet)
        {
            if (!string.Equals(sheet.Name, SheetName, StringComparison.CurrentCultureIgnoreCase) || !sheet.IsVisible)
            {
                #pragma warning disable CS0618
                LoadDataFilterOptions = None;
                return;
            }

            LoadDataFilterOptions = LoadDataFilterOptions.All;

        }
    }
}
