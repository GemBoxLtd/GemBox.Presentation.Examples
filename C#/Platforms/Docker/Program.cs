using System.IO;
using GemBox.Presentation;

class Program
{
    static void Main()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Create new presentation.
        var presentation = new PresentationDocument();

        // Create new slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Blank);

        // Add sample text.
        var shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 120, 120, 250, 150);
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Black));
        shape.Format.Fill.SetSolid(Color.FromHsl(0, 0, 240));

        var run = shape.Text.AddParagraph().AddRun("Lorem Ipsum");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Red));
        run.Format.Bold = true;

        run = shape.Text.AddParagraph().AddRun("Lorem Ipsum");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Green));
        run.Format.Italic = true;

        run = shape.Text.AddParagraph().AddRun("Lorem Ipsum");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Blue));
        run.Format.UnderlineStyle = UnderlineStyle.Single;

        // Add sample image.
        using (var stream = File.OpenRead("Dices.png"))
            slide.Content.AddPicture(PictureContentType.Png, stream, 480, 200, 240, 180);

        // Save presentation in PPTX and PDF format.
        presentation.Save("output.pptx");
        presentation.Save("output.pdf");
    }
}