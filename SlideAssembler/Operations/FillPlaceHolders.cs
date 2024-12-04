using ShapeCrawler;
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
                foreach (var paragraph in textFrame.Paragraphs)
                {
                    // if the paragraph doesn't contain a placeholder, move on
                    if (!PlaceholderRegex().IsMatch(paragraph.Text))
                    {
                        continue;
                    }

                    // Otherwise, we need to combine all portions that have the same formatting
                    // because sometimes Powerpoint splits the text into multiple portions,
                    // e.g. "*This* is a {{placeholder}}" might lead to the following portions:

                    // *This*
                    // is a
                    // {{
                    // placeholder
                    // }}

                    var combinablePortions = new List<IParagraphPortion>();
                    var enumerator = paragraph.Portions.GetEnumerator();

                    var hasMore = enumerator.MoveNext();

                    while (hasMore)
                    {
                        // Start with the current portion...

                        combinablePortions.Add(enumerator.Current);

                        // ...and collect all subsequent portions
                        // as long as they have the same formatting:

                        hasMore = enumerator.MoveNext();
                        while (hasMore)
                        {
                            if (HasEqualFormatting(enumerator.Current, combinablePortions.First()))
                            {
                                combinablePortions.Add(enumerator.Current);
                                hasMore = enumerator.MoveNext();
                            }
                            else
                            {
                                // We found a portion with different formatting.
                                // Try to combine all portions we collected so far
                                // and replace placeholders in the combined text.

                                ReplaceCombinedText(combinablePortions);

                                // Then start again to collect portions with equal formatting.
                                combinablePortions.Clear();
                                break;
                            }
                        }
                    }

                    // We reached the end of the paragraph.
                    // Combine all remaining portions and replace one final time.
                    ReplaceCombinedText(combinablePortions);
                }
            }
        }

        bool HasEqualFormatting(IParagraphPortion portion1, IParagraphPortion portion2)
        {
            return Compare(p => p.Font?.Color?.Type) &&
                   Compare(p => p.Font?.Color?.Hex) &&
                   Compare(p => p.Font?.Size) &&
                   Compare(p => p.Font?.OffsetEffect) &&
                   Compare(p => p.Font?.IsItalic) &&
                   Compare(p => p.Font?.EastAsianName) &&
                   Compare(p => p.Font?.LatinName) &&
                   Compare(p => p.Font?.IsBold) &&
                   Compare(p => p.Font?.Underline) &&
                   Compare(p => p.TextHighlightColor.Hex);

            bool Compare<T>(Func<IParagraphPortion, T> propertyAccessor)
            {
                var value1 = SafeGetValue(portion1);
                var value2 = SafeGetValue(portion2);

                return value1 is null
                    ? value2 is null
                    : value1.Equals(value2);

                T? SafeGetValue(IParagraphPortion portion)
                {
                    try
                    {
                        return propertyAccessor(portion);
                    }
                    catch
                    {
                        return default;
                    }
                }
            }
        }

        void ReplaceCombinedText(List<IParagraphPortion> combinablePortions)
        {
            var combinedText = string.Join(string.Empty, combinablePortions.Select(p => p.Text));
            var newText = ReplacePlaceholders(combinedText, context.ThrowOnError);
            if (newText != combinedText)
            {
                combinablePortions.First().Text = newText;
                foreach (var portion in combinablePortions.Skip(1))
                {
                    portion.Text = string.Empty;
                }
            };
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
                if (!throwOnError)
                {
                    return null;
                }
                else
                {
                    throw new InvalidDataException($"Missing data for placeholder: '{placeholder}'");
                }
            }

            var propertyInfo = currentObject.GetType().GetProperty(property.Trim());

            if (propertyInfo == null)
            {
                if (!throwOnError)
                {
                    return null;
                }
                else
                {
                    throw new InvalidDataException($"Missing data for placeholder: '{placeholder}'");
                }
            }

            currentObject = propertyInfo.GetValue(currentObject);
        }

        return currentObject;
    }

}
