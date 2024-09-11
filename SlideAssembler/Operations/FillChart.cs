using ShapeCrawler;
using SlideAssembler;


public class FillChart : IPresentationOperation
{
    private readonly object data;
    private bool ignoreMissingData;

    public FillChart(object data, bool ignoreMissingData)
    {
        this.data = data;
        this.ignoreMissingData = ignoreMissingData;
    }

    public void Apply(Presentation presentation)
    {
        throw new NotImplementedException();
    }
}

