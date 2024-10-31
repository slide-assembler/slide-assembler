using ShapeCrawler;
using SlideAssembler;

public partial class FillChart(string name, params Series[] seriesList) : IPresentationOperation
{
    public void Apply(PresentationContext context)
    {
        bool chartFound = false;

        foreach (var slide in context.Presentation.Slides)
        {
            var chart = slide.Shapes.OfType<IChart>().FirstOrDefault(c => c.Name == name);

            if (chart is not null)
            {
                chartFound = true;
                CompleteChart(context, chart);
            }
        }

        if (!chartFound && context.ThrowOnError)
        {
            throw new InvalidDataException($"Chart '{name}' not found.");
        }
    }

    public void CompleteChart(PresentationContext context, IChart chart)
    {
        foreach (var series in seriesList)
        {
            var chartSeries = chart.SeriesList.FirstOrDefault(s => s.Name == series.Name);

            if (chartSeries is null)
            {
                if (context.ThrowOnError)
                {
                    throw new InvalidDataException($"Series '{series.Name} in Chart '{name}' not found.");
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
