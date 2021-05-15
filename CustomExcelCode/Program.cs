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

    class Program
    {
        static async Task Main(string[] args)
        {
            string PROJECT_PATH = @"C:\Test-Projects\CustomExcelCode\CustomExcelCode";

            #region Write to Excel file

            // Set EPPlus to non commercial usage
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Person> samplePeople = Person.GetSamplePeople();


            // Demo 1: Saves the peopleData into an excel file
            FileInfo file1 = new(Path.Join(PROJECT_PATH, "myTestExcel.xlsx"));
            DeleteFileIfExists(file1);
            await CreateExcelFileAsync(file1, samplePeople);

            // Demo 2: Open a template excel. Save the peopleData with the current styles and save a new copy.
            FileInfo templatePath = new(Path.Join(PROJECT_PATH, "ExcelTemplate.xlsx"));
            FileInfo newExel = new(Path.Join(PROJECT_PATH, "ExcelCreatedFromTemplate.xlsx"));
            DeleteFileIfExists(newExel);
            await CreateExcelFromTemplate(newExel, templatePath, samplePeople);

            // MoreFunctionUtils();
            #endregion


            #region Read from Excel file
            FileInfo fileToRead = new(Path.Join(PROJECT_PATH, "myTestExcel.xlsx"));
            List<Person> peopleFromExcel = await LoadFromExcelFileAsync<Person>(fileToRead);


            #endregion



        }

        private static async Task<List<T>> LoadFromExcelFileAsync<T>(FileInfo file, int worksheetIndex = 0, int startingRow = 1)
        {
            // This process has to be custom process for fle types.
            // Since you have to figure out where each cell goes in the class model.
            List<T> result = new();
            PropertyInfo[] propertyInfo = typeof(T).GetProperties();

            using var package = new ExcelPackage(file);

            // This loads the file into memory
            await package.LoadAsync(file);

            // Choose the worksheet to load from
            var worksheet = package.Workbook.Worksheets[worksheetIndex];

            bool continueReadingFlag = true;
            while (continueReadingFlag)
            {
                bool atLeastOneCellHadValue = false;
                var typeInstance = Activator.CreateInstance(typeof(T));

                for (int i = 0; i < propertyInfo.Length; i++)
                {
                    object cellValue = worksheet.Cells[startingRow, i].Value;
                    if (string.IsNullOrEmpty(cellValue?.ToString()))
                    {
                        propertyInfo[i].SetValue(typeInstance, cellValue);
                        atLeastOneCellHadValue = true;
                    }
                }


                if (atLeastOneCellHadValue == false)
                {
                    continueReadingFlag = false;
                }

                result.Add((T)typeInstance);
            }



            return result;
        }

        private static async Task CreateExcelFileAsync<T>(FileInfo file, List<T> data)
        {
            // We create a package variable which is a wrapper to working with excel files.
            using ExcelPackage package = new(file);

            // Adds a worksheet to the excel file
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Main Worksheet");

            // Select a range to put the data into
            ExcelRange cellsRange = worksheet.Cells["A1"];

            // put data into cells range
            cellsRange.LoadFromCollection(data, true);

            // auto fit columns
            cellsRange.AutoFitColumns();

            // Save the changes in the excel file
            await package.SaveAsync();
        }

        private static async Task CreateExcelFromTemplate<T>(FileInfo file, FileInfo templatePath, List<T> data)
        {
            // Create a package from a template file
            using ExcelPackage package = new(file, templatePath);

            // Adds a worksheet to the excel file
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // Select the cells that have styles and we want to copy
            // ExcelStyle styles = worksheet.Cells["B1:B3"].Style;

            // We need to work with the cells one by one to apply the styles individually.
            char initialCol = 'A';
            int initialRow = 2;
            var properties = typeof(T).GetProperties();
            for (int i = 0; i < data.Count; i++)
            {
                int row = initialRow + i;
                for (int j = 0; j < properties.Length; j++)
                {
                    char col = (char)(initialCol + j);
                    // The cell to copy the styles from
                    ExcelRange templateCell = worksheet.Cells[$"{(char)(initialCol + j)}2"];


                    var cell = worksheet.Cells[$"{col}{row}"];
                    cell.LoadFromText(properties[j].GetValue(data[i])?.ToString());
                    // cell.Style.Fill = styles.Fill;
                    // cell.Style.Font = styles.Font;
                    // cell.Style.Border = styles.Border;
                    cell.StyleID = templateCell.StyleID;

                }
            }

            await package.SaveAsync();

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
