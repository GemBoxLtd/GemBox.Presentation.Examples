using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Set the directory path where the component will look for additional font files.
        // The "MyFonts" targets the subdirectory in the current directory, so besides the installed fonts,
        // the component will be able to use the fonts within the specified directory.
        FontSettings.FontsBaseDirectory = "MyFonts";

        var presentation = new PresentationDocument();

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
