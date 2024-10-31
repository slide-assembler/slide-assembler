using SlideAssembler;

namespace TestSlideAssembler
{
    [TestClass]
    public class UnitTestsFillChart
    {
        public FileStream template = File.OpenRead("Template.pptx");
        public FileStream templateForMultipleSeries = File.OpenRead("./verificationfiles/Template.pptx");

        [TestMethod]
        public void ThrowOnErrorIsFalse_ShouldNotThrowException()
        {
            var seriesWithMissingData = new Series("series1", []); // No data
            var fillChart = new FillChart("TestChart", seriesWithMissingData);

            Presentation.Load(template, throwOnError: false)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void ThrowOnErrorIsTrue_ShouldThrowException()
        {
            var seriesWithMissingData = new Series("series1", []); // No data
            var fillChart = new FillChart("TestChart", seriesWithMissingData);

            Presentation.Load(template, throwOnError: true)
                        .Apply(fillChart);
        }

        [TestMethod]
        [DataRow(true)]
        [DataRow(false)]
        [ExpectedException(typeof(NullReferenceException))]
        public void NullReferencesTest_singleSeries_RegardlessOfThrowOnError(bool throwOnError)
        {
            Series nullSeries = null;
            var fillChart = new FillChart("MesswertDiagramm", nullSeries);

            Presentation.Load(template, throwOnError)
                        .Apply(fillChart);
        }

        [TestMethod]
        [DataRow(true)]
        [DataRow(false)]
        [ExpectedException(typeof(NullReferenceException))]
        public void NullReferencesTest_MultipleSeries_RegardlessOfThrowOnError(bool throwOnError)
        {
            Series[] seriesList = [new Series("Datenreihe 1", [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart("LineChart", seriesList);

            Presentation.Load(templateForMultipleSeries, throwOnError: true)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void NullReferencesTest_NoSeriesName_ThrowOnErrorIsTrue()
        {
            Series[] seriesList = [new Series(null, [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart("LineChart", seriesList);

            Presentation.Load(templateForMultipleSeries, throwOnError: true)
                        .Apply(fillChart);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void NullReferencesTest_NoChartName_ThrowOnErrorIsTrue()
        {
            Series[] seriesList = [new Series("Series 1", [0.83, 0.1, 0.72, 0.15]), null, new Series("Datenreihe 2", [0.88, 0.8, 0.62, 0.75])];
            var fillChart = new FillChart(null, seriesList);

            Presentation.Load(templateForMultipleSeries, throwOnError: true)
                        .Apply(fillChart);
        }

        private void GeneratePresentation(FileStream template, double[] values, FileStream output)
        {
            Presentation.Load(template)
                        .Apply(new FillChart("MesswertDiagramm", new Series("Messwerte", values)))
                        .Save(output);
        }
    }
}
