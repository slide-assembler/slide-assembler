namespace TestSlideAssembler
{
    [TestClass]
    public class UnitTestsFillChart
    {
        [TestMethod]
        public void StandardTest()
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
        private void GeneratePresentation(FileStream template, dynamic data, FileStream output)
        {
            SlideAssembler.SlideAssembler.Load(template)
                        .Apply(new FillChart(data, false))
                        .Save(output);
        }
    }


}
