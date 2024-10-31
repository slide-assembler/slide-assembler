using System.Data;

namespace SlideAssembler.Tests;

[TestClass]
public class SetWidthUnitTests
{
    [TestMethod]
    [ExpectedException(typeof(InvalidDataException))]
    [DataRow(-1)]
    [DataRow(-0.1)]
    public void DecimalNotValidTest(double decimalnumber)
    {
        new SetWidth("shape", (Decimal)decimalnumber);
    }

    [TestMethod]
    [DataRow(0)]
    [DataRow(123)]
    public void DecimalValidTest(double decimalnumber)
    {
        new SetWidth("shape", (Decimal)decimalnumber);
    }
}
