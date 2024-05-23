using GemBox.Presentation;
using GemBox.Presentation.Media;
using System.IO;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create first presentation slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create first shape.
        var shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 2, 2, 8, 3, LengthUnit.Centimeter);

        // Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Create a paragraph.
        var paragraph = shape.Text.AddParagraph();

        // Add and format paragraph plain text.
        var run = paragraph.AddRun("Powered by ");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Add and format paragraph hyperlink text.
        run = paragraph.AddRun("GemBox");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed));
        run.Format.Action.Click.Set(ActionType.HyperlinkToWebPage, "http://www.gemboxsoftware.com/");

        // Create second presentation slide.
        slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create first shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 2, 2, 4, 3, LengthUnit.Centimeter);

        // Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Set text run.
        run = shape.Text.AddParagraph().AddRun("Play Sound");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed));

        // Set "play sound" action on created shape.
        var action = shape.Action.Click;
        using (var stream = File.OpenRead("Applause.wav"))
            action.PlaySound(stream, "applause.wav");

        action.Set(ActionType.None);
        action.Highlight = true;

        // Create second shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 7, 2, 4, 3, LengthUnit.Centimeter);

        // Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Set text run.
        run = shape.Text.AddParagraph().AddRun("Hyperlink To Previous Slide");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed));

        // Set "hyperlink to previous slide" action.
        shape.Action.Click.Set(ActionType.HyperlinkToPreviousSlide);

        // Create third shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 12, 2, 4, 3, LengthUnit.Centimeter);

        // Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Set text run.
        run = shape.Text.AddParagraph().AddRun("Hyperlink To Next Slide");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed));

        // Set "hyperlink to next slide" action.
        shape.Action.Click.Set(ActionType.HyperlinkToNextSlide);

        // Create forth shape.
        shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 17, 2, 4, 3, LengthUnit.Centimeter);

        // Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Set text run.
        run = shape.Text.AddParagraph().AddRun("Hyperlink To Web Page");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed));

        // Set "hyperlink to web page" action.
        shape.Action.Click.Set(ActionType.HyperlinkToWebPage, "http://www.gemboxsoftware.com/");

        // Create third presentation slide.
        slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        presentation.Save("Hyperlinks.pptx");
    }
}
