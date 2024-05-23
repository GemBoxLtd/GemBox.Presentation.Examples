using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("RightToLeft.pptx");

        var slide = presentation.Slides[0];

        var shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter);

        // Create a new right-to-left paragraph.
        var paragraph = shape.Text.AddParagraph();
        paragraph.Format.RightToLeft = true;
        paragraph.Format.Alignment = HorizontalAlignment.Right;
        var run = paragraph.AddRun("هذا ثمّة أمّا العالم، أم, السادس مواقعها");
        run.Format.Size = Length.From(28, LengthUnit.Point);

        presentation.Save("RightToLeft.pdf");
    }
}
