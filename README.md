# Slide Assembler

## Project Description
Slide Assembler is a .NET library designed to provide a fluent API for modifying PowerPoint presentations. It allows users to apply various operations to PowerPoint slides, such as filling charts, replacing placeholders, or modifying objects such as tables.

## Installation
To install Slide Assembler, you can use the NuGet package manager. Run the following command in the Package Manager Console:
```
Install-Package SlideAssembler
```

## Usage
Here are some examples of how to use the Slide Assembler library:

### Loading a Presentation
```csharp
using SlideAssembler;
using System.IO;

var stream = File.OpenRead("path/to/presentation.pptx");
var presentation = Presentation.Load(stream);
```

### Applying Operations
You can apply various operations to the presentation using the fluent API. For example, to fill a chart and replace placeholders:
```csharp
var data = new { Name = "John Doe", Age = 30 };
var series = new Series("Series1", new double[] { 1.0, 2.0, 3.0 });

presentation
    .Apply(new FillPlaceholders(data))
    .Apply(new FillChart("ChartName", series));
```

### Saving the Presentation
```csharp
using (var outputStream = new FileStream("path/to/output.pptx", FileMode.Create, FileAccess.Write))
{
    presentation.Save(outputStream);
}
```

## Build and Test
To build and test the project, you can use the provided GitHub Actions workflow. The workflow is defined in the `.github/workflows/dotnet.yml` file. It will build the project, run the tests, and upload the verification files and NuGet package as artifacts.

To run the tests locally, you can use the following command:
```
dotnet test
```
