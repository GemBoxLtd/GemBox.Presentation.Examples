using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Sets the base directory path where component looks for fonts.
        FontSettings.FontsBaseDirectory = ".";

        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        var textBox = slide.Content.AddTextBox(
            ShapeGeometryType.Rectangle, 2, 2, 8, 8, LengthUnit.Centimeter);

        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        var run = textBox.AddParagraph().AddRun(
            "Shows how to use private fonts with GemBox.Presentation component.");

        run.Format.Font = "Almonte Snow";
        run.Format.Size = Length.From(16, LengthUnit.Point);
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black));

        presentation.Save("Private Fonts.pdf");
    }
}