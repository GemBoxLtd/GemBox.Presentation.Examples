using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create new text box.
        var textBox = slide.Content.AddTextBox(
            ShapeGeometryType.RoundedRectangle, 2, 2, 10, 20, LengthUnit.Centimeter);

        // Create new paragraph.
        var paragraph = textBox.AddParagraph();

        // Create new run.
        var run = paragraph.AddRun("All caps: ");
        run = paragraph.AddRun("Capital letters");
        run.Format.Caps = CapsType.All;

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Bold: ");
        run = paragraph.AddRun("Bold text");
        run.Format.Bold = true;

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Italic: ");
        run = paragraph.AddRun("Italic text");
        run.Format.Italic = true;

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Underline: ");
        run = paragraph.AddRun("Single underline text");
        run.Format.UnderlineStyle = UnderlineStyle.Single;

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Font size: ");
        run = paragraph.AddRun("Font size is 14 points");
        run.Format.Size = Length.From(14, LengthUnit.Point);

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Strikethrough: ");
        run = paragraph.AddRun("Some text");
        run.Format.Strikethrough = StrikethroughType.Single;

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Double strikethrough: ");
        run = paragraph.AddRun("Some text");
        run.Format.Strikethrough = StrikethroughType.Double;

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Font color: ");
        run = paragraph.AddRun("Red text");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Red));

        paragraph.AddLineBreak();

        run = paragraph.AddRun("Font name: ");
        run = paragraph.AddRun("Arial Black");
        run.Format.Font = "Arial Black";

        presentation.Save("Character Formatting.pptx");
    }
}