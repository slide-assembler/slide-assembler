using ShapeCrawler;
using SlideAssembler;
using System.Text.RegularExpressions;


public partial class FillPlaceholders : IPresentationOperation
{
    private readonly object data;

    // Regular expression to find placeholders {{Name:Format}}
    [GeneratedRegex(@"{{(.*?)(:(.*?))?}}", RegexOptions.None)]
    private static partial Regex PlaceholderRegex();

    public FillPlaceholders(object data)
    {
        this.data = data;
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

    private string ReplacePlaceholders(string text, object data)
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
                if (!string.IsNullOrEmpty(format))
                {
                    double val = Convert.ToDouble(value);
                    formattedValue = val.ToString(format.Trim());
                }
                else
                {
                    formattedValue = value.ToString();
                }

                // replace the placeholder with the formatted value  
                text = text.Replace(match.Value, formattedValue);

            }
        }

        return text;
    }

    private object GetDataValue(object data, string placeholder)
    {

        var properties = placeholder.Split('.');
        object currentObject = data;


        foreach (var property in properties)
        {
            if (currentObject == null) return null;

            var propertyInfo = currentObject.GetType().GetProperty(property.Trim());
            if (propertyInfo == null) return null;

            currentObject = propertyInfo.GetValue(currentObject);
        }

        return currentObject;
    }

}