using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.Style;

namespace CustomExcelCode
{

    public record Person(int Id, string FirstName, string LastName)
    {
        public Person() : this(default, default, default)
        {
        }

        public static List<Person> GetSamplePeople()
        {
            List<Person> people = new();
            for (int i = 0; i < 10; i++)
            {
                people.Add(new(1, "Nadar", "Alpenidze"));
                people.Add(new(2, "Daria", "Liukin"));
                people.Add(new(3, "Yovel", "Gavrieli"));
                people.Add(new(4, "David", "Korochik"));
                people.Add(new(5, "Inna", "Korochik"));
            }
            return people;
        }
    }

    public class Program
    {
        public const string PROJECT_PATH = @"C:\Test-Projects\CustomExcelCode\CustomExcelCode";
        static async Task Main(string[] args)
        {

            

            #region Write to Excel file

            // Set EPPlus to non commercial usage
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Person> samplePeople = Person.GetSamplePeople();

            // Instantiate my service
            ExcelService excelService = new();


            // Demo 1: Saves the peopleData into an excel file
            FileInfo file1 = new(Path.Join(PROJECT_PATH, "myTestExcel.xlsx"));
            DeleteFileIfExists(file1);
            await excelService.WriteFileAsync(file1, samplePeople);

            // Demo 2: Open a template excel. Save the peopleData with the current styles and save a new copy.
            FileInfo templatePath = new(Path.Join(PROJECT_PATH, "ExcelTemplate.xlsx"));
            FileInfo newExel = new(Path.Join(PROJECT_PATH, "ExcelCreatedFromTemplate.xlsx"));
            DeleteFileIfExists(newExel);
            await excelService.WriteFileFromTemplateAsync(
                newExel,
                templatePath,
                samplePeople, 
                0,
                1,
                2,
                2
                );

            // MoreFunctionUtils();
            #endregion


            #region Read from Excel file
            FileInfo fileToRead = new(Path.Join(PROJECT_PATH, "myTestExcel.xlsx"));
            var people = await excelService.ReadFileAsync<Person>(fileToRead, startingRow: 2);


            #endregion



        }

      

        private static void DeleteFileIfExists(FileInfo file)
        {
            if (File.Exists(file.FullName))
            {
                File.Delete(file.FullName);
            }
        }


        /// <summary>
        /// This function is just a demo of some of the functionality we can use;
        /// </summary>
        private static void MoreFunctionUtils()
        {
            // This function
            using ExcelPackage package = new(new FileInfo("somefilewhere"));

            // Adds a worksheet to the excel file
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Main Worksheet");

            // Merge cells -> Merges the cells like you would do manually
            worksheet.Cells["A1:C1"].Merge = true;

            // Manipulate a whole column
            worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Changes all the column to be aligned horizontally to the center

            // Sets the height for all the cells in the row
            worksheet.Row(1).Height = 20;

            // Manipulates the styles of the font and puts to bold
            worksheet.Cells["B1:C1"].Style.Font.Bold = true;

        }
    }
}
