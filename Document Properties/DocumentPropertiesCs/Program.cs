using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        var slide = presentation.Slides[0];

        slide.Content.Drawings.Clear();

        // Create "Built-in document properties" text box.
        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 0.5, 0.5, 12, 10, LengthUnit.Centimeter);
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkBlue));

        var paragraph = textBox.AddParagraph();
        paragraph.Format.Alignment = HorizontalAlignment.Left;

        var run = paragraph.AddRun("Built-in document properties:");
        run.Format.Bold = true;

        paragraph.AddLineBreak();

        foreach (var docProp in presentation.DocumentProperties.BuiltIn)
        {
            paragraph.AddRun(string.Format("{0}: {1}", docProp.Key, docProp.Value));
            paragraph.AddLineBreak();
        }

        // Create "Custom document properties" text box.
        textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 14, 0.5, 12, 10, LengthUnit.Centimeter);
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkBlue));

        paragraph = textBox.AddParagraph();
        paragraph.Format.Alignment = HorizontalAlignment.Left;

        run = paragraph.AddRun("Custom document properties:");
        run.Format.Bold = true;

        paragraph.AddLineBreak();

        foreach (var docProp in presentation.DocumentProperties.Custom)
        {
            paragraph.AddRun(string.Format("{0}: {1} (Type: {2})", docProp.Key, docProp.Value, docProp.Value.GetType()));
            paragraph.AddLineBreak();
        }

        presentation.Save("Document Properties.pptx");
    }
}