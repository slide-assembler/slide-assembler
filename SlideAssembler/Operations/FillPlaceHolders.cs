using ShapeCrawler;
using SlideAssembler;
using System.Text.RegularExpressions;

public partial class FillPlaceholders : IPresentationOperation
{
    private readonly object data;
    private bool ignoreMissingData;

    // Regular expression to find placeholders {{Name:Format}}
    [GeneratedRegex(@"{{(.*?)(:(.*?))?}}", RegexOptions.None)]
    private static partial Regex PlaceholderRegex();

    public FillPlaceholders(object data, bool ignoreMissingData = false)
    {
        this.data = data;
        this.ignoreMissingData = ignoreMissingData;
    }

    public void Apply(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            foreach (var textFrame in slide.TextFrames())
            {
                var text = textFrame.Text;
                var newText = ReplacePlaceholders(text, data);
                textFrame.Text = newText;
            }
        }
    }

    public string ReplacePlaceholders(string text, object data)
    {
        var matches = PlaceholderRegex().Matches(text);

        foreach (Match match in matches)
        {
            var placeholder = match.Groups[1].Value;
            var format = match.Groups[3].Value;

            // Get the value from the data object
            var value = GetDataValue(data, placeholder);

            if (value != null)
            {
                // Apply formatting if specified
                string formattedValue;
                if (!string.IsNullOrEmpty(format) && value is IFormattable formattable)
                {
                    formattedValue = formattable.ToString(format.Trim(), null);
                }
                else
                {
                    formattedValue = value?.ToString() ?? string.Empty;
                }

                // replace the placeholder with the formatted value  
                text = text.Replace(match.Value, formattedValue);

            }
        }

        return text;
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
