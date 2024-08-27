using Microsoft.VisualStudio.TestPlatform.Utilities;
using SlideAssembler;

namespace TestSlideAssembler
{
    [TestClass]
    public class SetWidthUnitTests
    {
        [TestMethod]
        public void StandartTest()
        {
            var values = new[] { 82, 88, 64, 79, 31 };

            // Platzhalter
            var data = new
            {
                Titel = $"Messwerte vom {DateTime.Now:d}",
                Benutzer = new
                {
                    Vorname = "Philipp",
                    Nachname = "Kunnert"
                },
                Minimum = values.Min(),
                Mittelwert = values.Average(),
                Maximum = values.Max(),
                Werte = values
            };

            using var template = File.OpenRead("Template.pptx");
            using var output = new FileStream("Output.pptx", FileMode.Create, FileAccess.ReadWrite);

            SlideAssembler.SlideAssembler slideAssembler = SlideAssembler.SlideAssembler.Load(template);

            slideAssembler = slideAssembler.Apply(new FillPlaceHolders(data));
            slideAssembler.Apply(new SetWidth("MittelwertRechteck", (Decimal)data.Mittelwert))
                .Apply(new SetWidth("MaximumRechteck", (Decimal)data.Maximum))
                .Apply(new SetWidth("MinimumRechteck", (Decimal)data.Minimum))
                .Save(output);
        }

        [TestMethod]
        public void NullPlaceholderTest()
        {
            var values = new[] { 82, 88, 64, 79, 31 };

            // Platzhalter
            var data = new
            {
                Titel = $"Messwerte vom {DateTime.Now:d}",
                Benutzer = new
                {
                    Vorname = "Philipp",
                    Nachname = "Kunnert"
                },
                Minimum = values.Min(),
                Mittelwert = values.Average(),
                Maximum = values.Max(),
                Werte = values
            };

            using var template = File.OpenRead("Template.pptx");
            using var output = new FileStream("Output.pptx", FileMode.Create, FileAccess.ReadWrite);

            SlideAssembler.SlideAssembler slideAssembler = SlideAssembler.SlideAssembler.Load(template);

            slideAssembler = slideAssembler.Apply(new FillPlaceHolders(data));
            slideAssembler.Apply(new SetWidth("MittelwertRechteck", (Decimal)data.Mittelwert))
                .Apply(new SetWidth(null, (Decimal)data.Maximum))
                .Apply(new SetWidth(null, (Decimal)data.Minimum))
                .Save(output);
        }
    }
}