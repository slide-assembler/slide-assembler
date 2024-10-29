using ShapeCrawler;
using SlideAssembler;


public partial class ReplaceImage : IPresentationOperation
{
    private readonly Object data;
    private bool ignoreMissingData;

    public ReplaceImage(Object data, bool ignoreMissingData = false)
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

        if (property != null && property.PropertyType == typeof(Stream))
        {
            Stream stream = (Stream)property.GetValue(data);
            image.Image.Update(stream);
        }
    }
}

