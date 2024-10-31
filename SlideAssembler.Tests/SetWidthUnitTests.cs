using System.Data;

namespace SlideAssembler.Tests;

[TestClass]
public class SetWidthUnitTests
{
    [TestMethod]
    [ExpectedException(typeof(ArgumentOutOfRangeException))]
    [DataRow(-1)]
    [DataRow(-0.1)]
    public void WidthNotValidTest(double width)
    {
        new SetWidth("shape", (decimal)width);
    }

    [TestMethod]
    [DataRow(0)]
    [DataRow(1.5)]
    [DataRow(123)]
    public void WidthValidTest(double width)
    {
        new SetWidth("shape", (decimal)width);
    }
}
