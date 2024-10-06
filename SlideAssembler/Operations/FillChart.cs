
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
        this.seriesList = [singleSeries];
        this.ignoreMissingData = ignoreMissingData;
    }

    public void Apply(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape is IChart && shape.Name == chartTitle)
                {
                    CompleteChart((IChart)shape);
                }
            }
        }
    }

    private void CompleteChart(IChart chart)
    {
        for (int i = 0; i < seriesList.Length; i++)
        {
            var series = seriesList[i];
            var chartSeries = chart.SeriesList[i];
            if (chartSeries.Name == series.name)
            {
                for (int pointIndex = 0; pointIndex < series.values.Length; pointIndex++)
                {
                    if (pointIndex < chartSeries.Points.Count)
                    {
                        chartSeries.Points[pointIndex].Value = (double)series.values[pointIndex];
                    }
                }
            }
        }

    }

}

