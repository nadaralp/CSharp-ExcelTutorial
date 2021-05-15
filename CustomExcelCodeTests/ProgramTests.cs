using System;
using System.IO;
using System.Threading.Tasks;
using CustomExcelCode;
using FluentAssertions;
using Xunit;

namespace CustomExcelCodeTests
{
    public class ProgramTests : IDisposable
    {
        private readonly string _filePath;

        public ProgramTests()
        {
            _filePath = Path.Join(Program.PROJECT_PATH, "test.xlsx");
        }

        [Fact]
        public async Task Integration_CanCreateExcelFile_AndReadDataFromIt()
        {
            // Arrange
            ExcelService excelService = new();
            FileInfo file = new(_filePath);
            var sampleData = Person.GetSamplePeople();

            // Act
            await excelService.WriteFileAsync(file, sampleData, printHeaders: true);

            // start row is 2 since we print the headers
            var peopleData = await excelService.ReadFileAsync<Person>(file, startingRow: 2);

            // Assert
            peopleData.Should().HaveCountGreaterThan(0);
        }


        // Global teardown
        public void Dispose()
        {
            if (File.Exists(_filePath))
            {
                File.Delete(_filePath);
            }
        }
    }
}