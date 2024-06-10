using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load PowerPoint presentation; preservation feature is enabled by default.
        var presentation = PresentationDocument.Load("Preservation.pptx");
        var slide = presentation.Slides[0];

        // Modify the presentation.
        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 80, 300, 300, 80, LengthUnit.Point);
        textBox.AddParagraph().AddRun("You can preserve unsupported features when modifying a presentation!");

        // Save PowerPoint presentation to and output file of the same format together with
        // preserved information (unsupported features) from the input file.
        presentation.Save("PreservedOutput.pptx");
    }
}
