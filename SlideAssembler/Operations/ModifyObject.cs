using ShapeCrawler;
using SlideAssembler;



public class ModifyObject(string name, Action<IShape> action) : IPrestationOperation
{
    public void Apply(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape.Name.Equals(name))
                {
                    action(shape);
                }
            }
        }
    }
}

