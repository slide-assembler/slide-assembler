using ShapeCrawler;
using SlideAssembler;


public class FillChart : IPresentationOperation
{
    private readonly object data;
    private bool ignoreMissingData;

    public FillChart(object data, bool ignoreMissingData = false)
    {
        this.data = data;
        this.ignoreMissingData = ignoreMissingData;
    }

    public void Apply(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape is IChart)
                {
                    CompleteChart((IChart)shape);
                }
            }
        }
    }

    private void CompleteChart(IChart chart)
    {
        foreach (var series in chart.SeriesList)
        {
            for (int i = 0; i < series.Points.Count; i++)
            {
                var point = series.Points[i];
                string name = chart.Categories[i].Name;

                var value = GetDataValue(data, name);
                if (value == null)
                {
                    value = 0.0;
                }
                point.Value = (double)value;
            }
        }
    }

    public object? GetDataValue(object data, string placeholder)
    {

        var properties = placeholder.Split('.');
        object? currentObject = data;


        foreach (var property in properties)
        {
            if (currentObject == null)
            {
                if (ignoreMissingData) return null;
                throw new InvalidDataException("Data cant be null or empty!");
            }

            var propertyInfo = currentObject.GetType().GetProperty(property.Trim());

            if (propertyInfo == null && ignoreMissingData) return null;
            if (propertyInfo == null) throw new InvalidDataException("Missing Data!");

            currentObject = propertyInfo.GetValue(currentObject);
        }

        return currentObject;
    }

}

