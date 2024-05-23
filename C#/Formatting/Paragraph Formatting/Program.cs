using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create new text box.
        var textBox = slide.Content.AddTextBox(
            ShapeGeometryType.RoundedRectangle, 2, 2, 10, 4, LengthUnit.Centimeter);

        // Create new paragraph.
        var paragraph = textBox.AddParagraph();

        // Set paragraph text.
        paragraph.AddRun("This paragraph has the following properties: alignment is justify, after spacing is 100% of the text size, before spacing is 250% of the text size, line spacing is 200% of the text size.");

        // Set selected paragraph format.
        var format = paragraph.Format;
        format.Alignment = HorizontalAlignment.Justify;
        format.SpacingAfter = TextSpacing.Single;
        format.SpacingBefore = TextSpacing.Multiple(2.5);
        format.SpacingLine = TextSpacing.Double;

        // Create new paragraph.
        paragraph = textBox.AddParagraph();

        // Set paragraph text.
        paragraph.AddRun("This paragraph has the following properties: alignment is left, indentation before text is 15 points and first line indentation is 25 points.");

        // Set selected paragraph format.
        paragraph.Format.Alignment = HorizontalAlignment.Left;
        paragraph.Format.IndentationBeforeText = Length.From(15, LengthUnit.Point);
        paragraph.Format.IndentationSpecial = Length.From(25, LengthUnit.Point);

        presentation.Save("Paragraph Formatting.pptx");
    }
}
