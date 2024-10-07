using ShapeCrawler;

namespace SlideAssembler;


public class SlideAssembler
{
    private readonly Presentation presentation;

    private SlideAssembler(Presentation presentation) { this.presentation = presentation; }


    public static SlideAssembler Load(Stream stream) // load presantion from stream
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        try
        {
            return new SlideAssembler(new Presentation(stream));
        }
        catch (InvalidDataException ex)
        {
            throw ex;
        }

    }

    public void Save(Stream stream) // save prestation in stream 
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

    public SlideAssembler Apply(params IPresentationOperation[] operations) // applys changes and get the pdatet Prestation
    {
        foreach (var operation in operations)
        {
            operation.Apply(this.presentation);
        }
        return this;
    }

}

