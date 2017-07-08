using System;
using System.IO;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = new PresentationDocument();

        string pathToResources = "Resources";

        // Create new presentation slide.
        Slide slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create first picture from resource data.
        Picture picture = null;
        using (var stream = File.OpenRead(Path.Combine(pathToResources, "Dices.png")))
            picture = slide.Content.AddPicture(PictureContentType.Png, stream, 2, 2, 6, 5, LengthUnit.Centimeter);

        // Create "rounded rectangle" shape.
        Shape shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 10, 2, 8, 5, LengthUnit.Centimeter);

        // Fill shape with picture content.
        PictureFillFormat fillFormat = shape.Format.Fill.SetPicture(picture.Fill.Data.Content);

        // Set the offset of the edges of the stretched picture fill.
        fillFormat.StretchLeft = 0.1;
        fillFormat.StretchRight = 0.4;
        fillFormat.StretchTop = 0.1;
        fillFormat.StretchBottom = 0.4;

        // Get shape outline format.
        LineFormat lineFormat = shape.Format.Outline;

        // Set shape red outline.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red));
        lineFormat.Width = Length.From(0.2, LengthUnit.Centimeter);

        presentation.Save("Pictures.pptx");
    }
}
