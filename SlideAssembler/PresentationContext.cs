namespace SlideAssembler;

public record PresentationContext(
    ShapeCrawlerPresentation Presentation,
    bool ThrowOnError = false);
