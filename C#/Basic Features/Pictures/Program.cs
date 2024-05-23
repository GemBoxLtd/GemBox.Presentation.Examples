using GemBox.Presentation;
using System.IO;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new presentation slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create first picture from resource data.
        Picture picture = null;
        using (var stream = File.OpenRead("Dices.png"))
            picture = slide.Content.AddPicture(PictureContentType.Png, stream, 2, 2, 6, 5, LengthUnit.Centimeter);

        // Create "rounded rectangle" shape.
        var shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 10, 2, 8, 5, LengthUnit.Centimeter);

        // Fill shape with picture content.
        var fillFormat = shape.Format.Fill.SetPicture(picture.Fill.Data.Content);

        // Set the offset of the edges of the stretched picture fill.
        fillFormat.StretchLeft = 0.1;
        fillFormat.StretchRight = 0.4;
        fillFormat.StretchTop = 0.1;
        fillFormat.StretchBottom = 0.4;

        // Get shape outline format.
        var lineFormat = shape.Format.Outline;

        // Set shape red outline.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red));
        lineFormat.Width = Length.From(0.2, LengthUnit.Centimeter);

        // Create second picture from SVG resource.
        using (var stream = File.OpenRead("Graphics1.svg"))
            picture = slide.Content.AddPicture(PictureContentType.Svg, stream, 2, 8, 6, 3, LengthUnit.Centimeter);

        presentation.Save("Pictures.pptx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Input Pictures.pptx");
        var slide = presentation.Slides[0];

        // Get all pictures from first slide.
        var pictures = slide.Content.Drawings.All().OfType<Picture>();

        // Get first picture data.
        Picture picture = pictures.First();
        PictureContent pictureContent = picture.Fill.Data;

        // Export picture data to image file.
        using (var fileStream = File.Create($"Output.{pictureContent.ContentType}"))
        using (var pictureStream = pictureContent.Content.Open())
            pictureStream.CopyTo(fileStream);
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Input Pictures.pptx");
        var slide = presentation.Slides[0];

        // Get all pictures from first slide.
        var pictures = slide.Content.Drawings.All().OfType<Picture>();

        // Replace pictures data with image file.
        foreach (var picture in pictures)
            using (var fileStream = File.OpenRead("Jellyfish.jpg"))
                picture.Fill.SetData(PictureContentType.Jpeg, fileStream);

        presentation.Save("Updated Pictures.pptx");
    }
}
