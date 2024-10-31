using System.Data;
using System.Globalization;

namespace SlideAssembler.Tests;

[TestClass]
public class UnitTestsFillPlaceHolders
{
    [DataTestMethod]
    [DataRow("de-AT", "xyz 5,32")]
    [DataRow("en-US", "xyz 5.32")]
    public void TestDoubleFormat_ReplacePlaceholders(string culture, string expected)
    {
        CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = new CultureInfo(culture);

        var data = new { MyProperty = 5.321 };
        var operation = new FillPlaceholders(data);
        var result = operation.ReplacePlaceholders("xyz {{MyProperty:N2}}", throwOnError: true);

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

        var operation = new FillPlaceholders(data);
        var result = operation.ReplacePlaceholders("Hello my name is {{name}}, i live in {{address}} and i earn {{salary:N2}}", throwOnError: false);

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
        var result = operation.ReplacePlaceholders(placeholders, throwOnError: false);

        Assert.AreEqual(expected, result);
    }

    [TestMethod]
    [DataRow(true)]
    [DataRow(false)]
    public void IgnoreMissingData_GetDataValue(bool throwOnError)
    {
        string placeholder = "{{name}}, {{age}}";
        var data = new
        {
            name = "Max Mustermann",
        };

        var operation = new FillPlaceholders(data);

        if (throwOnError)
        {
            Assert.ThrowsException<InvalidDataException>(() => operation.GetDataValue(placeholder, throwOnError));
        }
        else
        {
            Assert.AreEqual(null, operation.GetDataValue(placeholder, throwOnError));
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
        Presentation.Load(template, throwOnError: true)
                    .Apply(new FillPlaceholders(data))
                    .Save(output);
    }
}