using SlideAssembler;

var fileName = "Output_Datenpraesentation.pptx";

var placeholderData = new
{
    House = "Seniorenpension Musterhaus",
    Date = "Dezember 2023",
    Weight = new
    {
        Median = 68,
        Minimum = 30,
        Maximum = 104
    }
};

using var template = File.OpenRead("Template_Datenpraesentation.pptx");
using var output = File.Create(fileName);
{
    Presentation.Load(template)
        .Apply(new FillPlaceholders(placeholderData))
        .Apply(new FillChart(
            "Häufigkeitsverteilung", [
                new Series("Datenreihe 1", [1, 5, 8]),
                new ("Datenreihe 2", [17, 9, 10])]))
        .Save(output);
}