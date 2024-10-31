using System.Data;
using System.Globalization;
using SlideAssembler;

namespace TestSlideAssembler
{
    [TestClass]
    public class UnitTestsFillPlaceHolders
    {
        [TestMethod]
        public void StandartTest()
        {
            var values = new[] { 0.82, 0.88, 0.64, 0.79, 0.31 };

            var data = new
            {
                Titel = $"Messwerte vom {DateTime.Now:d}",
                Benutzer = new
                {
                    Vorname = "Max",
                    Nachname = "Mustermann"
                },
                Minimum = values.Min(),
                Mittelwert = values.Average(),
                Maximum = values.Max(),
                Werte = values
            };

            using var template = File.OpenRead("Template.pptx");
            using var output = new FileStream("Output.pptx", FileMode.Create, FileAccess.ReadWrite);
            GeneratePresentation(template, data, output);
        }

        [DataTestMethod]
        [DataRow("de-AT", "xyz 5,32")]
        [DataRow("en-US", "xyz 5.32")]
        public void TestDoubleFormat_ReplacePlaceholders(string culture, string expected)
        {
            CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = new CultureInfo(culture);

            var data = new { MyProperty = 5.321 };
            var operation = new FillPlaceholders(data);
            var result = operation.ReplacePlaceholders("xyz {{MyProperty:N2}}", data);

            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void MissingDataTest_ReplacePlaceholders()
        {
            CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = new CultureInfo("de-AT");

            var data = new
            {
                name = "Max Mustermann",
                salary = 521.23
            };

            var operation = new FillPlaceholders(data, true);
            var result = operation.ReplacePlaceholders("Hello my name is {{name}}, i live in {{address}} and i earn {{salary:N2}}", data);

            Assert.AreEqual("Hello my name is Max Mustermann, i live in {{address}} and i earn 521,23", result);
        }

        [TestMethod]
        [DataRow("{{name}}, {{age}}, {{salary}}$", "Max Mustermann, 18, 718$")]
        [DataRow("", "")]
        [DataRow("{{name}}, {{salary}}$", "Max Mustermann, 718$")]
        public void MissingPlaceholdersTest(string placeholders, string expected)
        {
            var data = new
            {
                name = "Max Mustermann",
                age = 18,
                salary = 718
            };

            var operation = new FillPlaceholders(data);
            var result = operation.ReplacePlaceholders(placeholders, data);

            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void IgnoreMissingData_GetDataValue(bool ignoreMissingData)
        {
            string placeholder = "{{name}}, {{age}}";
            var data = new
            {
                name = "Max Mustermann",
            };

            var operation = new FillPlaceholders(data, ignoreMissingData);

            if (!ignoreMissingData)
            {
                Assert.ThrowsException<InvalidDataException>(() => operation.GetDataValue(data, placeholder));
            }
            else
            {
                Assert.AreEqual(null, operation.GetDataValue(data, placeholder));
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void NullObjectTest()
        {
            using var template = File.OpenRead("Template.pptx");
            using var output = new FileStream("Output_null_data.pptx", FileMode.Create, FileAccess.ReadWrite);
            GeneratePresentation(template, null, output);
        }

        [TestMethod]
        [ExpectedException(typeof(FileLoadException))]
        public void NotAPowerpointTest()
        {
            using var template = File.OpenRead("Not_a_PowerPoint.txt");
            using var output = new FileStream("Output_not_a_Powerpoint.pptx", FileMode.Create, FileAccess.ReadWrite);
            GeneratePresentation(template, null, output);
        }

        private void GeneratePresentation(FileStream template, dynamic data, FileStream output)
        {
            Presentation.Load(template)
                        .Apply(new FillPlaceholders(data, false))
                        .Save(output);
        }
    }
}