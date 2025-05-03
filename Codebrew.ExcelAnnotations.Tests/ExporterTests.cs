using ClosedXML.Excel;
using Codebrew.ExcelAnnotations.Attributes;

namespace Codebrew.ExcelAnnotations.Tests
{
    [ExcelSheet("Users")]
    public class UserDto
    {
        [ExcelColumn("Name")]
        public string Name { get; set; }

        [ExcelColumn("Age")]
        public int Age { get; set; }

        [ExcelColumn("Birth Date")]
        public DateTime BirthDate { get; set; }
    }

    public class ExporterTests
    {
        [Test]
        public void Export_CreatesValidExcelWithCorrectHeadersAndData()
        {
            // Arrange
            var users = new List<UserDto>
            {
                new UserDto { Name = "Alice", Age = 30, BirthDate = new DateTime(1994, 1, 1) },
                new UserDto { Name = "Bob", Age = 25, BirthDate = new DateTime(1999, 5, 20) }
            };

            // Act
            Stream stream = Exporter.Export(users);
            stream.Position = 0;
            var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet("Users");

            // Assert - headers
            Assert.AreEqual("Name", worksheet.Cell(1, 1).Value);
            Assert.AreEqual("Age", worksheet.Cell(1, 2).Value);
            Assert.AreEqual("Birth Date", worksheet.Cell(1, 3).Value);

            // Assert - data rows
            Assert.AreEqual("Alice", worksheet.Cell(2, 1).Value);
            Assert.AreEqual(30, worksheet.Cell(2, 2).GetValue<int>());
            Assert.AreEqual(new DateTime(1994, 1, 1), worksheet.Cell(2, 3).GetDateTime());

            Assert.AreEqual("Bob", worksheet.Cell(3, 1).Value);
            Assert.AreEqual(25, worksheet.Cell(3, 2).GetValue<int>());
            Assert.AreEqual(new DateTime(1999, 5, 20), worksheet.Cell(3, 3).GetDateTime());
        }
    }
}
