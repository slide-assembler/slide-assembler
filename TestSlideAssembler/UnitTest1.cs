namespace TestSlideAssembler
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
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

            using var template = File.OpenRead("Template.pptx");
            using var output = new FileStream("Output.pptx", FileMode.Create, FileAccess.ReadWrite);

            SlideAssembler.SlideAssembler slideAssembler = SlideAssembler.SlideAssembler.Load(template);

            slideAssembler = slideAssembler.Apply(new FillPlaceHolders(data));
            slideAssembler.Apply(new SetWidth("MittelwertRechteck", (Decimal)data.Mittelwert))
                .Apply(new SetWidth("MaximumRechteck", (Decimal)data.Maximum))
                .Apply(new SetWidth("MaximumRechteck", (Decimal)data.Minimum))
                .Save(output);



        }
    }
}