using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace CustomExcelCode
{
    public interface IExcelService
    {
        Task WriteFileAsync<T>(FileInfo file,
            IList<T> data,
            string sheetName = "Sheet1",
            bool printHeaders = true
        );


        Task WriteFileFromTemplateAsync<T>(
            FileInfo file,
            FileInfo templateFile,
            IList<T> data,
            int worksheetIndex = 0,
            int initialColumnToWriteTo = 1,
            int initialRowToWriteTo = 1,
            int templateRow = 1
        );


        Task<IList<T>> ReadFileAsync<T>(FileInfo file, int worksheetIndex = 0, int startingRow = 1);
    }

    public class ExcelService : IExcelService
    {
        public ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task WriteFileAsync<T>(FileInfo file, IList<T> data, string sheetName = "Sheet1", bool printHeaders = true)
        {
            // We create a package variable which is a wrapper to working with excel files.
            using ExcelPackage package = new(file);

            // Adds a worksheet to the excel file
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

            // Select a range to put the data into
            ExcelRange cellsRange = worksheet.Cells["A1"];

            // put data into cells range
            cellsRange.LoadFromCollection(data, printHeaders);

            // auto fit columns
            cellsRange.AutoFitColumns();

            // Save the changes in the excel file
            await package.SaveAsync();
        }

        public async Task WriteFileFromTemplateAsync<T>(
            FileInfo file,
            FileInfo templateFile,
            IList<T> data,
            int worksheetIndex = 0,
            int initialColumnToWriteTo = 1,
            int initialRowToWriteTo = 1,
            int templateRow = 1
            )
        {
            // Create a package from a template file
            using ExcelPackage package = new(file, templateFile);

            // Adds a worksheet to the excel file
            ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetIndex];

            // Select the cells that have styles and we want to copy
            // We need to work with the cells one by one to apply the styles individually.

            var properties = typeof(T).GetProperties();
            for (int i = 0; i < data.Count; i++)
            {
                int row = initialRowToWriteTo + i;
                for (int j = 0; j < properties.Length; j++)
                {
                    int col = initialColumnToWriteTo + j;
                    // The cell to copy the styles from
                    ExcelRange templateCell = worksheet.Cells[$"{(char)(initialColumnToWriteTo + j + 64 /*64 = 'A' - 1*/)}{templateRow}"];


                    var cell = worksheet.Cells[row, col];
                    cell.LoadFromText(properties[j].GetValue(data[i])?.ToString());
                    // cell.Style.Fill = styles.Fill;
                    // cell.Style.Font = styles.Font;
                    // cell.Style.Border = styles.Border;
                    cell.StyleID = templateCell.StyleID;

                }
            }

            await package.SaveAsync();
        }


        public async Task<IList<T>> ReadFileAsync<T>(FileInfo file, int worksheetIndex = 0, int startingRow = 1)
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

            while (true)
            {
                bool atLeastOneCellHadValue = false;
                var typeInstance = Activator.CreateInstance(typeof(T));

                for (int i = 0; i < propertyInfo.Length; i++)
                {
                    object cellValue = worksheet.Cells[startingRow, i + 1].Value;
                    if (string.IsNullOrEmpty(cellValue?.ToString()) == false)
                    {
                        dynamic value = Convert.ChangeType(cellValue, propertyInfo[i].PropertyType);
                        propertyInfo[i].SetValue(typeInstance, value);
                        atLeastOneCellHadValue = true;
                    }
                }


                if (atLeastOneCellHadValue == false)
                {
                    break;
                }

                result.Add((T)typeInstance);
                startingRow += 1;
            }



            return result;
        }
    }
}
