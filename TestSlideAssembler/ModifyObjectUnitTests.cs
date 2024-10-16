namespace TestSlideAssembler
{
    [TestClass]
    public class ModifyObjectUnitTests
    {
        FileStream template = File.OpenRead("Template.pptx");

        [TestMethod]
        public void IgnoreMissingData_isTrue_SHouldNotThrowException()
        {
            SlideAssembler.SlideAssembler slideAssembler = SlideAssembler.SlideAssembler.Load(template)
                .Apply(new ModifyObject("NotValidObject", o =>
                {
                    o.TextFrame.Text = "text1";
                    o.Width = (Decimal)82;
                    o.Height = 40;
                }, true));
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void IgnoreMissingData_isfalse_ShouldThrowException()
        {
            SlideAssembler.SlideAssembler slideAssembler = SlideAssembler.SlideAssembler.Load(template)
                .Apply(new ModifyObject("NotValidObject", o =>
                {
                    o.TextFrame.Text = "text1";
                    o.Width = (Decimal)82;
                    o.Height = 40;
                }, false));
        }

        [TestMethod]
        public void StandartTest()
        {
            var values = new[] { 0.82, 0.88, 0.64, 0.79, 0.31 };

            // Platzhalter
            var data = new
            {
                Titel = $"Messwerte vom {DateTime.Now:d}",
                Benutzer = new
                {
                    Vorname = "Philipp",
                    Nachname = "Kunnert"
                },
                Minimum = values.Min(),
                Mittelwert = values.Average(),
                Maximum = values.Max(),
                Werte = values
            };

            using var output = new FileStream("Output.pptx", FileMode.Create, FileAccess.ReadWrite);

            SlideAssembler.SlideAssembler slideAssembler = SlideAssembler.SlideAssembler.Load(template);

            slideAssembler = slideAssembler.Apply(new FillPlaceholders(data));
            slideAssembler
                .Apply(new SetWidth("MaximumRechteck", (Decimal)data.Maximum))
                .Apply(new SetWidth("MaximumRechteck", (Decimal)data.Minimum));

            slideAssembler.Apply(new ModifyObject("MittelwertRechteck", o =>
            {
                o.TextFrame.Text = data.Mittelwert.ToString("N2");
                o.Width = (Decimal)data.Mittelwert;
                o.Height = 40;
            }));

            slideAssembler.Save(output);
        }
    }
}