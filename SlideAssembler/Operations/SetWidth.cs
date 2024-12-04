using SlideAssembler;


public class SetWidth : IPresentationOperation
{
    private readonly string name;
    private readonly decimal width;

    public SetWidth(string name, decimal width)
    {
        if (width < 0.0m)
        {
            throw new ArgumentOutOfRangeException("Width has to be >= 0", nameof(width));
        }

        this.name = name;
        this.width = width;
    }

    public void Apply(PresentationContext context)
    {
        var shapeFound = false;

        foreach (var slide in context.Presentation.Slides)
        {
            var shape = slide.Shapes.FirstOrDefault(s => s.Name == name);
            if (shape is not null)
            {
                shape.Width = width;
                shapeFound = true;
            }
        }

        if (!shapeFound && context.ThrowOnError)
        {
            throw new InvalidDataException($"Shape '{name}' not found.");
        }
    }
}

