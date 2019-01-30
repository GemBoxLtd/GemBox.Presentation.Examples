using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        var shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 2, 2, 8, 4, LengthUnit.Centimeter);
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.DarkBlue));

        var run = shape.Text.AddParagraph().AddRun("This sample shows how to write or save a new PowerPoint file with GemBox.Presentation.");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.White));

        presentation.Save("Writing.pptx");
    }
}