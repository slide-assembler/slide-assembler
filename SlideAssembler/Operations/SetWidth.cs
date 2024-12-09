using ShapeCrawler;
using SlideAssembler;
using SlideAssembler.Operations;


public class SetWidth : NamedShapeOperation<IShape>
{
    private readonly string name;
    private readonly decimal width;

    public SetWidth(string name, decimal width) : base(name)
    {
        if (width < 0.0m)
        {
            throw new ArgumentOutOfRangeException("Width has to be >= 0", nameof(width));
        }

        this.name = name;
        this.width = width;
    }

    protected override void Apply(PresentationContext context, IShape shape)
    {
        shape.Width = width;

        if (shape == null && context.ThrowOnError)
        {
            throw new InvalidDataException($"Shape '{name}' not found.");
        }
    }


}

