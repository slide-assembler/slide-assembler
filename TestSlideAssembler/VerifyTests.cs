using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Imaging;
using System.IO;
using Spire.Presentation;
using System.Security.AccessControl;
using Newtonsoft.Json;
using System.Text.Json;

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
            GeneratePowerpointWithAllFeatures();
            await VerifyAllSlides();
        }

        private async Task VerifyAllSlides()
        {
            if (!Directory.Exists(verificationDirectory))
            {
                Directory.CreateDirectory(verificationDirectory);
            }

            using (Spire.Presentation.Presentation presentation = new Presentation())
            {
                presentation.LoadFromFile("verify-tests-Output.pptx", FileFormat.Auto);

                foreach (ISlide slide in presentation.Slides)
                {
                    var image = slide.SaveAsImage();

                    using (var imageStream = new MemoryStream())
                    {
                        image.Save(imageStream, ImageFormat.Png);
                        imageStream.Position = 0;

                        await Verify(imageStream, "png")
                            .UseDirectory(verificationDirectory)
                            .UseFileName("slide-" + slide.SlideNumber);       //expecting all slide-xy.verified.png pictures are already there
                    }
                }
            }
        }

        private void GeneratePowerpointWithAllFeatures()    //uses SlideAssembler to create a Presentation with all its features 
        {
            using (var template = File.OpenRead(Path.Combine(verificationDirectory, powerpointTemplate)))
            {
                using (var output = new FileStream("verify-tests-Output.pptx", FileMode.Create, FileAccess.ReadWrite))
                {
                    //====================================================Data for Powerpoint generation (has to match the template)====================================================

                    var values = new[] { 0.82, 0.88, 0.64, 0.79, 0.31 };

                    var data = new
                    {
                        Titel = $"Messwerte vom {DateTime.Now:d}",
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
                                .Save(output);
                }
            }
        }
    }
}
