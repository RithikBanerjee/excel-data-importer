
![Poster](/Assets/posters/Importer.png)

# Excel Data Importer

&emsp;&emsp; [Excel Data Importer](/ExcelDataImporter) is a class library which you can use to import excel data into your own software or database in just few fractions of a second. It runs on .Net framework 4.6.1 and c# verion 6.0  and such a fast excel processing is done by aspose cells and c data libraries of .Net. This is the ultimate software management tool for excel to tally or any other software data import as per the given schema.<br />
&emsp;&emsp; The data importer needs just the excel file path and the schema file path to which schema it would parse the excel data. Then it validates the excel schema with the schema provided, if error in schema it returns message as provided. Then it validates the excel data with the schema provided and parses the valid data and invalid data as a result for the excel import.

# Tables of Content

- [Workbook Schema](#Workbook-Schema)
- [Light Cell Data Handlers](#Light-Cell-Data-Handlers)
- [Data Builder](#Data-Builder)
- [Data Importers](#Data-Importers)
- [FAQ](#faq)


## Workbook Schema 
&emsp;&emsp; [Workbook Schema](/ExcelDataImporter/Model/WorkbookSchema.cs) is the excel schema structure or the model in which the Json schema provided is parsed and used to process. The Json schema has many useful properties as you can set the names of the column to find in the excel, regex pattern for every column, message on invalid regex for any row, default values for empty column and many more. And [here](/DemoApp/JsonSchema/DemoSchema.json) is an example of the json schema you need to provide before you try import excel. <br />
&emsp;&emsp; The complete structure of the library depends on this model and if the provided schema is incorrect then the output data won't be as expected. 

## Light Cell Data Handlers
&emsp;&emsp; This includes one abstract class which is needed to be inherited in case you need to import your excel data into your own class. Actually, this [base](/ExcelDataImporter/LightCellDataHandlers/BaseLightCellDataHandler.cs) class inherits the light cell data handler of Aspose Cell's which allows to iterate over the excel cells at an unimaginable rate as well as process every cell indiviually with the superclass methods StartEachRow, ProcessCellFurther and ProcessRowFurther as per the object you need.<br />
&emsp;&emsp; There are few are two demo light cell handlers created to give directions on how you can create your own object's light cell handler. One is with the traditional datatable which needs no modification and the other is with a [demo](/ExcelDataImporter/Model/DemoTable.cs) class.

## Data Builder
&emsp;&emsp; [Builder](/ExcelDataImporter/Builder/DataBuilder.cs) is a class file which is constant and needs no modification. It just uses C Data's queries to validate schema and prepare the sheet in order to read it easily and faster for the Aspose Cell's handler. Major responsibilities of this class files includes delete rows above header rows, delete blank rows, remove formatting etc.

## Data Importers
&emsp;&emsp; This is the public method which is called from the excel is imported and hence it is also needed to be modified by inheriting the abstract [base](/ExcelDataImporter/DataImporter/BaseDataImporter.cs) class which is responsible for adding row index, formating date, sorting the data according to the unique field in order to help map the those objects which have One-2-Many or Many-2-Many relation which their properties.

## FAQ

#### How to run the project?
A [Demo App](/DemoApp) is made to elastrate how to use excel importer with a [demo excel](/DemoApp/DemoExcel.xlsx).<br />
Open [Excel Data Importer](../../blob/master/ExcelDataImporter.sln) in your visual studio and Press 'F5'.

#### What's the minimum framework needed?
.Net Framework 4.6.1 & C# 6.0 <br />
All required dll's are given in [Builds](/Builds).

#### How is it so fast?
All thanks to C Data's fast schema validation and Aspose Cell's fast data validation.

please contribute bros?
