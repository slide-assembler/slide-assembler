//using GemBox.Presentation;

var path = "..\\..\\..\\..\\TestSlideAssembler\\bin\\debug\\net8.0\\verify-tests-Output.pptx";

//ComponentInfo.SetLicense("FREE-LIMITED-KEY");

//var presentation = PresentationDocument.Load(path);

//presentation.Slides.RemoveAt(0);
//presentation.Save("1.pdf", ImageSaveOptions.Pdf);


//var presentation = Syncfusion.Presentation.Presentation.Open(path);
//presentation.PresentationRenderer = new PresentationRenderer();
//using var stream = presentation.Slides.First().ConvertToImage(Syncfusion.Presentation.ExportImageFormat.Png);
//using var file = File.Open("1.png", FileMode.Create);
//await stream.CopyToAsync(file);


var presentation = FileFormat.Slides.Presentation.Open(path);