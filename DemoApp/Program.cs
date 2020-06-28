
using System;
using System.IO;
using ExcelDataImporter.DataImporter;

namespace DemoApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Choose path option:\n1.Have excel file path? \n2.Use default file path?");
            if (int.TryParse(Console.ReadLine(), out var pathOption))
            {
                if(Equals(pathOption, 1))
                {
                    Console.WriteLine("\nYour excel file path:");
                    ReadyToImport(Console.ReadLine());
                }
                if(Equals(pathOption, 2))
                    ReadyToImport(DirectoryPath("DemoExcel.xlsx"));
            }
            Console.ReadKey();
        }

        static void ReadyToImport(string excelPath)
        {
            //example for datatypes: datatable and object for demo purpose
            Console.WriteLine("Choose data option:\n1.Import in datatable? \n2.Import in collection of objects?");
            if (int.TryParse(Console.ReadLine(), out var dataOption))
            {
                var schemaPath = DirectoryPath(@"JsonSchema\DemoSchema.json");
                if (Equals(dataOption, 1))
                    Console.WriteLine(Import(new DemoDataTableImporter(excelPath, schemaPath)));

                if(Equals(dataOption, 2))
                    Console.WriteLine(Import(new DemoTableImporter(excelPath, schemaPath)));
            }
        }

        static string DirectoryPath(string folderDirectory)
        {
            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\")) + folderDirectory;
        }

        static string Import<T>(BaseDataImporter<T> dataImporter)
        {
            if (!dataImporter.ValidateSchema())
                return $"\n{dataImporter.Workbook.InvalidSchema.Rows[0]["WhyInvalid"]}";
            if (!dataImporter.ValidateData())
                return "\nInvalid Data";
            //put debugger here and check dataImporter.Workbook.Sheets[0].ValidData
            var yourExcelData = dataImporter.Workbook.Sheets[0].ValidData;
            return "\nImported Successfully";
        }
    }
}
