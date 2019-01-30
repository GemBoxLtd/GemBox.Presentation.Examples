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

        // Create new "rounded rectangle" shape.
        var shape = slide.Content.AddShape(
            ShapeGeometryType.RoundedRectangle, 2, 2, 5, 4, LengthUnit.Centimeter);

        // Get shape format.
        var format = shape.Format;

        // Get shape fill format.
        var fillFormat = format.Fill;

        // Set shape fill format as solid fill.
        fillFormat.SetSolid(Color.FromName(ColorName.DarkBlue));

        // Create new "rectangle" shape.
        shape = slide.Content.AddShape(
            ShapeGeometryType.Rectangle, 8, 2, 5, 4, LengthUnit.Centimeter);

        // Set shape fill format as solid fill.
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow));

        // Set shape outline format as solid fill.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green));

        // Create new "rounded rectangle" shape.
        shape = slide.Content.AddShape(
            ShapeGeometryType.RoundedRectangle, 14, 2, 5, 4, LengthUnit.Centimeter);

        // Set shape fill format as no fill.
        shape.Format.Fill.SetNone();

        // Get shape outline format.
        var lineFormat = shape.Format.Outline;

        // Set shape outline format as single solid red line.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red));
        lineFormat.DashType = LineDashType.Solid;
        lineFormat.Width = Length.From(0.8, LengthUnit.Centimeter);
        lineFormat.CompoundType = LineCompoundType.Single;

        presentation.Save("Shape Formatting.pptx");
    }
}