using SlideAssembler;


public class SetWidth : IPresentationOperation
{
    private readonly string name;
    private decimal data;

    public SetWidth(string name, decimal data)
    {
        this.name = name;
        if (data >= 0) this.data = data;
        else throw new InvalidDataException("Width of Shape has to be > 0");

    }
    public void Apply(PresentationContext context)
    {
        var shapeFound = false;

        foreach (var slide in context.Presentation.Slides)
        {
            var shape = slide.Shapes.FirstOrDefault(s => s.Name == name);
            if (shape is not null)
            {
                shape.Width = data;
                shapeFound = true;
            }
        }

        if (!shapeFound && context.ThrowOnError)
        {
            throw new InvalidDataException($"Shape '{name}' not found.");
        }
    }
}

