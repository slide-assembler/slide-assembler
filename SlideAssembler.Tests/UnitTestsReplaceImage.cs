using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ShapeCrawler;
using SlideAssembler;

namespace TestSlideAssembler
{
    [TestClass]
    public class UnitTestsReplaceImage
    {
        [TestMethod]
        public void Standarttest()
        {
            var imageData = new
            {
                Logo = File.OpenRead("Template.pptx")
            };
            var operation = new ReplaceImage(imageData);

            using var presentationStream = File.OpenRead("Template.pptx");

            SlideAssembler.SlideAssembler.Load(presentationStream).Apply(new ReplaceImage(imageData)).Save(new MemoryStream());
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void ReplaceImage_WithMissingData_ThrowsException()
        {
            var imageData = new
            {
                NotTheRightLogo = new MemoryStream(new byte[] { /* byte data for an image */ })
            };
            var operation = new ReplaceImage(imageData);

            using var presentationStream = File.OpenRead("Template.pptx");

            operation.Apply(new Presentation(presentationStream));
        }

        [TestMethod]
        public void ReplaceImage_WithMissingData_IgnoresException_WhenIgnoreMissingDataIsTrue()
        {
            var imageData = new
            {
                NotTheRightLogo = new MemoryStream(new byte[] { /* byte data for an image */ })
            };
            var operation = new ReplaceImage(imageData, true);

            using var presentationStream = File.OpenRead("Template.pptx");

            operation.Apply(new Presentation(presentationStream));
            Assert.IsTrue(true);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidDataException))]
        public void ReplaceImage_InvalidStream_ThrowsException()
        {
            var imageData = new
            {
                Logo = "This is not a stream"
            };
            var operation = new ReplaceImage(imageData);

            using var presentationStream = File.OpenRead("Template.pptx");

            operation.Apply(new Presentation(presentationStream));
        }
    }
}
