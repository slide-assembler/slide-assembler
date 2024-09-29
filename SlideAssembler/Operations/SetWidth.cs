using ShapeCrawler;
using SlideAssembler;


public class SetWidth : IPrestationOperation
{
    private readonly string name;
    private decimal data;

    public SetWidth(string name, decimal data)
    {
        this.name = name;
        if (data != null && data > 0) this.data = data;
        else throw new InvalidDataException("Width of Shape has to be > 0");
        
    }
    public void Apply(Presentation presentation)
    {
        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape.Name.Equals(name))
                {
                    shape.Width = data;
                    shape.TextFrame.Text = data.ToString();
                }
            }
        }
    }
}

