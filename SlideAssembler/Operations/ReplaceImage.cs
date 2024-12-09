using ShapeCrawler;
using SlideAssembler;


public partial class ReplaceImage : IPresentationOperation
{
    private readonly object data;
    private bool ignoreMissingData;

    public ReplaceImage(object data)
    {
        this.data = data;
    }
    public void Apply(PresentationContext context)
    {
        foreach (var slide in context.Presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape is IPicture)
                {
                    UpdateImage((IPicture)shape, context.ThrowOnError);
                }
            }
        }
    }

    private void UpdateImage(IPicture image, bool throwOnError = false)
    {
        var property = data.GetType().GetProperty(image.Name);

        if (property != null && typeof(Stream).IsAssignableFrom(property.PropertyType))
        {
            Stream stream = (Stream)property.GetValue(data);
            image.Image.Update(stream);
        }
        else if (throwOnError)
        {
            throw new InvalidDataException("Property is null or not a stream");
        }
    }
}

