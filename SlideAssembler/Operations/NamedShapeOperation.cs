using ShapeCrawler;

namespace SlideAssembler.Operations
{
    public abstract class NamedShapeOperation<TShape>(string name) : IPresentationOperation
        where TShape : IShape
    {
        public virtual void Apply(PresentationContext context)
        {
            foreach (var slide in context.Presentation.Slides)
            {
                var shape = slide.Shapes.OfType<TShape>().FirstOrDefault(c => c.Name == name);

                if (shape is not null)
                {
                    this.Apply(context, shape);
                    return;
                }
            }

            if (context.ThrowOnError)
            {
                throw new InvalidDataException($"'{name}' not found.");
            }
        }

        protected abstract void Apply(PresentationContext context, TShape shape);
    }
}
