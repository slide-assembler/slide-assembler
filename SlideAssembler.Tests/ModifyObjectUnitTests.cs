namespace SlideAssembler.Tests;

[TestClass]
public class ModifyObjectUnitTests
{
    FileStream template = File.OpenRead("Template.pptx");

    [TestMethod]
    public void ThrowOnError_IsFalse_ShouldNotThrowException()
    {
        Presentation.Load(template, false)
            .Apply(new ModifyObject("NotValidObject", o =>
            {
                o.TextBox.Text = "text1";
                o.Width = (Decimal)82;
                o.Height = 40;
            }));
    }

    [TestMethod]
    [ExpectedException(typeof(InvalidDataException))]
    public void ThrowOnError_IsTrue_ShouldThrowException()
    {
        Presentation.Load(template, true)
            .Apply(new ModifyObject("NotValidObject", o =>
            {
                o.TextBox.Text = "text1";
                o.Width = (Decimal)82;
                o.Height = 40;
            }));
    }
}