using ShapeCrawler;

namespace SlideAssembler;


class SlideAssembler
{
    private readonly Presentation presentation;

    private SlideAssembler(Presentation presentation) { this.presentation = presentation; }

    public static Presentation Load(Stream stream) // Loads presantion from Stream
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        try
        {
            return new Presentation(stream);
        }
        catch (InvalidDataException ex)
        {
            throw ex;
        }

    }

    public void Save(Stream stream) // Saves prestation in Stream 
    {
        if (stream == null) throw new ArgumentNullException("Stream can not be null.");
        try
        {
            presentation.SaveAs(stream);
        }
        catch (InvalidDataException ex)
        {
            throw ex;
        }

    }

    public SlideAssembler Apply(params IPrestationOperation[] operations) // Applys chnages and get the Updatet Prestation
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

