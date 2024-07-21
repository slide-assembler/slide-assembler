using ShapeCrawler;
using SlideAssembler;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

public class FillPlaceHolders : IPrestationOperation
{
    private readonly object data;

    public FillPlaceHolders(object data)
    {
        this.data = data;
    }

    public void Apply(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape is ITextFrame textFrame)
                {
                    var text = textFrame.Text;
                    var newText = ReplacePlaceholders(text, data);
                    textFrame.Text = newText;
                }
            }
        }
    }

    private string ReplacePlaceholders(string text, object data)
    {
        // Regular expression to find placeholders {{Name:Format}}
        var regex = new Regex(@"{{(.*?)(:(.*?))?}}");
        var matches = regex.Matches(text);

        foreach (Match match in matches)
        {
            var placeholder = match.Groups[1].Value; // Platzhaltername
            var format = match.Groups[3].Value; // Optionales Format ohne :

            // Get the value from the data object
            var value = GetDataValue(data, placeholder);

            if (value != null)
            {
                // Apply formatting if specified
                string formattedValue;
                if (!string.IsNullOrEmpty(format))
                {
                    formattedValue = string.Format($"{{0:{format}}}", value);
                }
                else
                {
                    formattedValue = value.ToString();
                }

                // Replace the placeholder with the formatted value
                text = text.Replace(match.Value, formattedValue);
            }
        }

        return text;
    }

    private object GetDataValue(object data, string placeholder)
    {
        // Use a stack to perform a depth-first search on all properties of the data object
        var stack = new Stack<object>();
        stack.Push(data);

        while (stack.Count > 0)
        {
            var current = stack.Pop();
            if (current == null)
                continue;

            var type = current.GetType();
            var propertyInfo = type.GetProperty(placeholder, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
            if (propertyInfo != null)
            {
                return propertyInfo.GetValue(current);
            }

            foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (prop.PropertyType.IsClass && prop.PropertyType != typeof(string))
                {
                    stack.Push(prop.GetValue(current));
                }
            }
        }

        return null;
    }
}
