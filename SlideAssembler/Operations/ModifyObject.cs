using ShapeCrawler;
using SlideAssembler;

public class ModifyObject(string name, Action<IShape> action) : IPresentationOperation
{
    public void Apply(PresentationContext context)
    {
        bool shapeFound = false;

        foreach (var slide in context.Presentation.Slides)
        {
            var shape = slide.Shapes.FirstOrDefault(s => s.Name == name);
            if (shape is not null)
            {
                action(shape);
                shapeFound = true;
            }
        }

        if (!shapeFound && context.ThrowOnError)
        {
            throw new InvalidDataException("Object '{name}' not found");
        }
    }
}

