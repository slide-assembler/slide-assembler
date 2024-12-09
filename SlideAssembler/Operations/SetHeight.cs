using ShapeCrawler;
using SlideAssembler;
using SlideAssembler.Operations;

public class SetHeight : NamedShapeOperation<IShape>
{
    private readonly string name;
    private readonly decimal height;

    public SetHeight(string name, decimal height) : base(name)
    {
        if (height < 0.0m)
        {
            throw new ArgumentOutOfRangeException("Height has to be >= 0", nameof(height));
        }

        this.name = name;
        this.height = height;
    }

    protected override void Apply(PresentationContext context, IShape shape)
    {
        shape.Height = height;
        if (shape == null && context.ThrowOnError)
        {
            throw new InvalidDataException($"Shape '{name}' not found.");
        }
    }
}

