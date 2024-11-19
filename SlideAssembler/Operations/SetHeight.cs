using SlideAssembler;

public class SetHeight : IPresentationOperation
{
    private readonly string name;
    private readonly decimal height;

    public SetHeight(string name, decimal height)
    {
        if (height < 0.0m)
        {
            throw new ArgumentOutOfRangeException("Height has to be >= 0", nameof(height));
        }

        this.name = name;
        this.height = height;
    }


    void IPresentationOperation.Apply(PresentationContext context)
    {
        var shapeFound = false;

        foreach (var slide in context.Presentation.Slides)
        {
            var shape = slide.Shapes.FirstOrDefault(s => s.Name == name);
            if (shape is not null)
            {
                shape.Height = height;
                shapeFound = true;
            }
        }

        if (!shapeFound && context.ThrowOnError)
        {
            throw new InvalidDataException($"Shape '{name}' not found.");
        }
    }
}

