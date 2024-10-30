using ShapeCrawler;
using SlideAssembler;


public partial class ReplaceImage : IPresentationOperation
{
    private readonly object data;
    private bool ignoreMissingData;

    public ReplaceImage(object data, bool ignoreMissingData = false)
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
                if (shape is IPicture)
                {
                    UpdateImage((IPicture)shape);
                }
            }
        }
    }

    private void UpdateImage(IPicture image)
    {
        var property = data.GetType().GetProperty(image.Name);

        if (property != null && typeof(Stream).IsAssignableFrom(property.PropertyType))
        {
            Stream stream = (Stream)property.GetValue(data);
            image.Image.Update(stream);
        }
        else if (!ignoreMissingData)
        {
            throw new InvalidDataException("Property is null or not a stream");
        }
    }
}

