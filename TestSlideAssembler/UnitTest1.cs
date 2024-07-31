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

            SlideAssembler.SlideAssembler.Load(template)
                        .Apply(new FillPlaceHolders(data))
                        .Save(output);
        }
    }
}