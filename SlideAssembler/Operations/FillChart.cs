using ShapeCrawler;
using SlideAssembler;

public partial class FillChart : IPresentationOperation
{
    private string chartTitle;
    private Series[] seriesList;
    private bool ignoreMissingData;

    public FillChart(string chartTitle, Series[] seriesList, bool ignoreMissingData = false)
    {
        this.chartTitle = chartTitle;
        this.seriesList = seriesList;
        this.ignoreMissingData = ignoreMissingData;
    }

    public FillChart(string chartTitle, Series singleSeries, bool ignoreMissingData = false)
    {
        this.chartTitle = chartTitle;
        this.seriesList = new Series[] { singleSeries };
        this.ignoreMissingData = ignoreMissingData;
    }

    public void Apply(Presentation presentation)
    {
        bool chartFound = false;

        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape is IChart chart && shape.Name == chartTitle)
                {
                    chartFound = true;
                    CompleteChart(chart);
                }
            }
        }

        if (!chartFound && !ignoreMissingData)
        {
            throw new InvalidDataException($"Chart mit dem Titel '{chartTitle}' wurde nicht gefunden.");
        }
    }

    public void CompleteChart(IChart chart)
    {
        bool seriesFound = false;

        for (int i = 0; i < seriesList.Length; i++)
        {
            if (seriesList[i] == null && !ignoreMissingData)
            {
                throw new NullReferenceException("Series cannot be null!");
            }
            else if (seriesList[i] == null)
            {
                i++;
            }
            else
            {
                var series = seriesList[i];

                var chartSeries = chart.SeriesList.FirstOrDefault(s => s.Name == series.name);

                if (chartSeries != null)
                {
                    seriesFound = true;
                    for (int pointIndex = 0; pointIndex < series.values.Length; pointIndex++)
                    {
                        if (pointIndex < chartSeries.Points.Count)
                        {
                            chartSeries.Points[pointIndex].Value = (double)series.values[pointIndex];
                        }
                    }
                }
            }

            if (!seriesFound && !ignoreMissingData)
            {
                throw new InvalidDataException($"Serie(n) in Chart '{chartTitle}' wurden nicht gefunden.");
            }
        }
    }
}
