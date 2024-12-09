using System.Globalization;
using System.Runtime.CompilerServices;
using ShapeCrawler;
using Syncfusion.PresentationRenderer;

namespace SlideAssembler.Tests;

[TestClass]
public class VerifyTests : VerifyBase
{
    private const string verificationDirectory = "verificationfiles"; //directory, where the template and all verification slide snapshots are located. (directory will be created automatically in the first run)

    [TestMethod]
    public async Task TestDifferentFormatting()
    {
        CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = new CultureInfo("de-AT");

        using var stream = new MemoryStream();
        using var template = File.OpenRead(Path.Combine(verificationDirectory, "Template_different_formating.pptx"));

        var data = new
        {
            Titel = "TestPräsentation",
            Projekt = "Slideassembler",
            Datum = "31.10.2024",
            Data1 = new
            {
                Headline = "Data 1",
                Date = "1.1.2001",
                Value = 20,
                Time = "17:30"
            },
            Data2 = new
            {
                Headline = "Data 2",
                Date = "2.1.2001",
                Value = 19,
                Time = "18:00"
            }
        };

        Presentation.Load(template)
                    .Apply(new FillPlaceholders(data))
                    .Save(stream);

        stream.Position = 0;

        await VerifyAllSlidesAsync(stream);
    }

    [TestMethod]
    public async Task TestAllFeatures()
    {
        CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = new CultureInfo("de-AT");

        using var stream = new MemoryStream();

        using var template = File.OpenRead(Path.Combine(verificationDirectory, "Template_all_features.pptx"));
        {
            var values = new[] { 82d, 88d, 64d, 79d, 31d };

            var data = new
            {
                Titel = $"Messwerte vom {new DateTime(2024, 8, 7):d}",
                Benutzer = new
                {
                    Vorname = "Max",
                    Nachname = "Mustermann"
                },
                Minimum = values.Min(),
                Mittelwert = values.Average(),
                Maximum = values.Max(),
                Werte = values
            };

            var seriesList = new[]
            {
                new Series("Datenreihe 1", values),
                new Series("Datenreihe 2", values.Select(v => v * 2).ToArray())
            };

            Presentation
                .Load(template)
                .Apply(new FillPlaceholders(data))
                .Apply(new FillChart("MesswertDiagramm", new Series("Werte", values)))
                .Apply(new SetWidth("MittelwertRechteck", (decimal)data.Mittelwert))
                .Apply(new SetWidth("MaximumRechteck", (decimal)data.Maximum))
                .Apply(new SetWidth("MinimumRechteck", (decimal)data.Minimum))
                .Apply(new SetHeight("MinimumRechteck", (decimal)data.Minimum))
                .Apply(new FillChart("LineChart", seriesList))
                .Apply(new ModifyObject<IShape>("MittelwertRechteck", o =>
                {
                    // TODO: Create extra object for modify object test
                    o.TextBox.Text = data.Mittelwert.ToString("N2");
                    o.Width = (Decimal)data.Mittelwert;
                    o.Height = 10;
                }))
                .Save(stream);

        }

        stream.Position = 0;

        await VerifyAllSlidesAsync(stream);
    }

    private async Task VerifyAllSlidesAsync(Stream stream, [CallerMemberName] string testName = default!)
    {
        if (!Directory.Exists(verificationDirectory))
        {
            Directory.CreateDirectory(verificationDirectory);
        }

        using var presentation = Syncfusion.Presentation.Presentation.Open(stream);
        presentation.PresentationRenderer = new PresentationRenderer();
        var slides = presentation.RenderAsImages(Syncfusion.Presentation.ExportImageFormat.Png);

        for (int i = 0; i < slides.Length; i++)
        {
            await Verify(slides[i], "png")
                .UseDirectory(verificationDirectory)
                .UseFileName($"{testName}-{i + 1}");
        }
    }
}
