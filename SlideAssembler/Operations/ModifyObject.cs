using ShapeCrawler;
using SlideAssembler;
using SlideAssembler.Operations;

public class ModifyObject<TShape>(string name, Action<TShape> action) : NamedShapeOperation<TShape>(name)
    where TShape : IShape
{
    protected override void Apply(PresentationContext context, TShape shape)
    {
        action(shape);
    }
}

