using ShapeCrawler;
using SlideAssembler;

public class ModifyObject : IPresentationOperation
{
    private readonly string name;
    private readonly Action<IShape> action;
    private readonly bool ignoreMissingData;

    public ModifyObject(string name, Action<IShape> action, bool ignoreMissingData = false)
    {
        this.name = name;
        this.action = action;
        this.ignoreMissingData = ignoreMissingData;
    }

    public void Apply(ShapeCrawlerPresentation presentation)
    {
        bool shapeFound = false;

        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape.Name.Equals(name))
                {
                    action(shape);
                    shapeFound = true;
                }
            }
        }

        if (!shapeFound && !ignoreMissingData)
        {
            throw new InvalidDataException("Object not found!");
        }
    }
}

