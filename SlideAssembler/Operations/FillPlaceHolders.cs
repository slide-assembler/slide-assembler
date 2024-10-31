using SlideAssembler;
using System.Text.RegularExpressions;

public partial class FillPlaceholders(object data) : IPresentationOperation
{
    // Regular expression to find placeholders {{Name[:Format]}}
    [GeneratedRegex(@"{{(.*?)(:(.*?))?}}", RegexOptions.None, 1000)]
    private static partial Regex PlaceholderRegex();

    public void Apply(PresentationContext context)
    {
        foreach (var slide in context.Presentation.Slides)
        {
            foreach (var textFrame in slide.TextFrames())
            {
                var text = textFrame.Text;
                var newText = ReplacePlaceholders(text, context.ThrowOnError);
                if (newText != textFrame.Text)
                {
                    textFrame.Text = newText;
                }
            }
        }
    }

    public string ReplacePlaceholders(string text, bool throwOnError)
    {
        var matches = PlaceholderRegex().Matches(text);

        foreach (Match match in matches)
        {
            var placeholder = match.Groups[1].Value;
            var format = match.Groups[3].Value;

            // Get the value from the data object
            var value = GetDataValue(placeholder, throwOnError);

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

    public object? GetDataValue(string placeholder, bool throwOnError)
    {

        var properties = placeholder.Split('.');
        object? currentObject = data;


        foreach (var property in properties)
        {
            if (currentObject == null)
            {
                if (!throwOnError) return null;
                throw new InvalidDataException("Data cant be null or empty!");
            }

            var propertyInfo = currentObject.GetType().GetProperty(property.Trim());

            if (propertyInfo == null && !throwOnError) return null;
            if (propertyInfo == null) throw new InvalidDataException("Missing Data!");

            currentObject = propertyInfo.GetValue(currentObject);
        }

        return currentObject;
    }

}
