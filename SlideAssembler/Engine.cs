using ShapeCrawler;

namespace SlideAssembler;


public class SlideAssembler
{
    private readonly Presentation presentation;

    private SlideAssembler(Presentation presentation) { this.presentation = presentation; }

    public static SlideAssembler Load(Stream stream) // load Presentation from stream
    {
        if (stream == null || !stream.CanRead)
            throw new ArgumentException("Stream is null or not readable.", nameof(stream));

        if (!IsPowerPointStream(stream))
            throw new FileLoadException("The File given is not a Powerpoint Presentation!");

        try
        {
            return new SlideAssembler(new Presentation(stream));
        }
        catch (InvalidDataException ex)
        {
            throw ex;
        }

    }

    public static bool IsPowerPointStream(Stream stream)
    {
        if (stream == null || !stream.CanRead)
            throw new ArgumentException("Stream is null or not readable.", nameof(stream));

        byte[] pptSignature = { 0xD0, 0xCF, 0x11, 0xE0 };
        byte[] pptxSignature = { 0x50, 0x4B, 0x03, 0x04 };

        // Read the first 4 bytes (signature)
        byte[] buffer = new byte[4];
        int bytesRead = stream.Read(buffer, 0, buffer.Length);

        stream.Position = 0;

        if (bytesRead < 4)
            return false; // Not enough data to determine file type

        if (buffer[0] == pptSignature[0] && buffer[1] == pptSignature[1] &&
            buffer[2] == pptSignature[2] && buffer[3] == pptSignature[3])
        {
            return true;
        }

        if (buffer[0] == pptxSignature[0] && buffer[1] == pptxSignature[1] &&
            buffer[2] == pptxSignature[2] && buffer[3] == pptxSignature[3])
        {
            return true;
        }

        return false;
    }

    public void Save(Stream stream) // save Presentation in stream 
    {
        if (stream == null) throw new ArgumentNullException("Stream can´t be null.");
        try
        {
            presentation.SaveAs(stream);
        }
        catch (InvalidDataException ex)
        {
            throw ex;
        }

    }

    public SlideAssembler Apply(params IPresentationOperation[] operations) // applies changes and get the updated Presentation
    {
        foreach (var operation in operations)
        {
            operation.Apply(this.presentation);
        }
        return this;
    }
}

