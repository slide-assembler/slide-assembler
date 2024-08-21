using System.Globalization;
using Syncfusion.PresentationRenderer;

namespace TestSlideAssembler
{
    [TestClass]
    public class VerifyTests : VerifyBase
    {

        //============================================================================Changeable Variables ============================================================================

        private string verificationDirectory = "verificationfiles"; //directory, where the template and all verification slide snapshots are located. (directory will be created automatically in the first run)

        const string powerpointTemplate = "Template.pptx"; //template where the presentation is created from and then verified (has to include all features)

        //=============================================================================================================================================================================

        [TestMethod]
        public async Task TestGeneratedPresentation()
        {
            CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = new CultureInfo("de-AT");

            using var stream = new MemoryStream();
            GeneratePowerpointWithAllFeatures(stream);
            stream.Position = 0;
            await VerifyAllSlides(stream);
        }

        private async Task VerifyAllSlides(Stream stream)
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
                    .UseFileName("slide-" + (i + 1));
            }
        }

        private void GeneratePowerpointWithAllFeatures(Stream stream)    //uses SlideAssembler to create a Presentation with all its features 
        {
            using (var template = File.OpenRead(Path.Combine(verificationDirectory, powerpointTemplate)))
            {
                //====================================================Data for Powerpoint generation (has to match the template)====================================================

                var values = new[] { 0.82, 0.88, 0.64, 0.79, 0.31 };

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

                //==================================================================================================================================================================

                SlideAssembler.SlideAssembler.Load(template)
                            .Apply(new FillPlaceholders(data))
                            .Save(stream);
            }
        }
    }
}
