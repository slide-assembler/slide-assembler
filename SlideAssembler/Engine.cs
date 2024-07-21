using ShapeCrawler;

namespace SlideAssembler;


class SlideAssembler
{
    private readonly Presentation presentation;

    private SlideAssembler(Presentation presentation) { this.presentation = presentation; }

    public static Presentation Load(Stream stream)
    { // Loads presantion from Stream

        return new Presentation(stream);
    }

    public void Save(Stream stream) // Saves prestation in Stream 
    {
        presentation.SaveAs(stream);
    }

    public SlideAssembler Apply(params IPrestationOperation[] operations)
    {
        foreach (var operation in operations)
        {
            operation.Apply(presentation);
        }
        return this;
    }

}

interface IPrestationOperation
{
    void Apply(Presentation presentation);
}

