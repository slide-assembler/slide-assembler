using SlideAssembler;

namespace TestSlideAssembler
{
    [TestClass]
    public class UnitTestsFillChart
    {
        public FileStream template = File.OpenRead("Template.pptx");
        public FileStream templateForMultipleSeries = File.OpenRead("./verificationfiles/Template.pptx");

        [TestMethod]
        public void IgnoreMissingDataIsTrue_ShouldNotThrowException()
        {
            var seriesWithMissingData = new Series("series1", []); // No data
            var fillChart = new FillChart("TestChart", new[] { seriesWithMissingData }, ignoreMissingData: true);

            Presentation.Load(template)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void IgnoreMissingDataIsFalse_ShouldThrowException()
        {
            var seriesWithMissingData = new Series("series1", []); // No data
            var fillChart = new FillChart("TestChart", new[] { seriesWithMissingData }); //ignoreMissingData is False if its untouched

            Presentation.Load(template)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void nullReferencesTest_singleSeries_IgnoreMissingDataIsFalse()
        {
            Series nullSeries = null;
            var fillChart = new FillChart("MesswertDiagramm", nullSeries, false);

            Presentation.Load(template)
                        .Apply(fillChart);
        }

        [TestMethod]
        public void nullReferencesTest_singleSeries_IgnoreMissingDataIsTrue()
        {
            Series nullSeries = null;
            var fillChart = new FillChart("MesswertDiagramm", nullSeries, true);

            Presentation.Load(template)
                        .Apply(fillChart);
        }


        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void nullReferencesTest_MultipleSeries_IgnoreMissingDataIsFalse()
        {
            Series[] seriesList = [new Series("Datenreihe 1", [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart("LineChart", seriesList, false);

            Presentation.Load(templateForMultipleSeries)
                        .Apply(fillChart);
        }

        [TestMethod]
        public void nullReferencesTest_MultipleSeries_IgnoreMissingDataIsTrue()
        {
            Series[] seriesList = [new Series("Datenreihe 1", [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart("LineChart", seriesList, true);

            Presentation.Load(templateForMultipleSeries)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void nullReferencesTest_NoSeriesName_IgnoreMissingDataIsfalse()
        {
            Series[] seriesList = [new Series(null, [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart("LineChart", seriesList, false);

            Presentation.Load(templateForMultipleSeries)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void nullReferencesTest_NoChartName_IgnoreMissingDataIsfalse()
        {
            Series[] seriesList = [new Series("Series 1", [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart(null, seriesList, false);

            Presentation.Load(templateForMultipleSeries)
                        .Apply(fillChart);
        }

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

            using var output = new FileStream("Output.pptx", FileMode.Create, FileAccess.ReadWrite);
            GeneratePresentation(template, values, output);


        }
        private void GeneratePresentation(FileStream template, double[] values, FileStream output)
        {
            Presentation.Load(template)
                        .Apply(new FillChart("MesswertDiagramm", new Series("Messwerte", values)))
                        .Save(output);
        }
    }


}
