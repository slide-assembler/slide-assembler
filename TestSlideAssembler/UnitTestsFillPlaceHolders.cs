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

        [TestMethod]
        public void MissingDataTest()
        {
            var values = new[] { 0.82, 0.88, 0.64};

            var data = new
            {
                Titel = $"Messwerte vom {DateTime.Now:d}",
                Benutzer = new
                {
                    Vorname = "Max"
                },
                Minimum = values.Min(),
                Werte = values
            };

            using var template = File.OpenRead("Template.pptx");
            using var output = new FileStream("Output_missing_Data.pptx", FileMode.Create, FileAccess.ReadWrite);
            GeneratePresentation(template, data, output);
        }

        [TestMethod]
        public void MissingPlaceholdersTest()
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

            using var template = File.OpenRead("Template_missing_Placeholders.pptx");
            using var output = new FileStream("Output_missing_Placeholders.pptx", FileMode.Create, FileAccess.ReadWrite);
            GeneratePresentation(template, data, output);
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
            SlideAssembler.SlideAssembler.Load(template)
                        .Apply(new FillPlaceholders(data))
                        .Save(output);
        }
    }
}