using ShapeCrawler;
using SlideAssembler;
using SlideAssembler.Operations;

public partial class FillChart(string name, params Series[] seriesList) : NamedShapeOperation<IChart>(name)
{
    protected override void Apply(PresentationContext context, IChart chart)
    {
        foreach (var series in seriesList)
        {
            var chartSeries = chart.SeriesList.FirstOrDefault(s => s.Name == series.Name);

            if (chartSeries is null)
            {
                if (context.ThrowOnError)
                {
                    throw new InvalidDataException($"Series '{series.Name} in chart '{name}' not found.");
                }
                else
                {
                    continue;
                }
            }

            for (int pointIndex = 0; pointIndex < series.Values.Length; pointIndex++)
            {
                if (pointIndex < chartSeries.Points.Count)
                {
                    chartSeries.Points[pointIndex].Value = (double)series.Values[pointIndex];
                }
            }
        }
    }
}
