using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new slide; will create "custom" layout slide and default master slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        slide.Content.AddShape(ShapeGeometryType.RectangularCallout, 30, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.AliceBlue));
        slide.Content.AddShape(ShapeGeometryType.RoundedRectangularCallout, 170, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.BlueViolet));
        slide.Content.AddShape(ShapeGeometryType.OvalCallout, 310, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.CadetBlue));
        slide.Content.AddShape(ShapeGeometryType.Pentagon, 450, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.CornflowerBlue));

        slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 30, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.DarkSeaGreen));
        slide.Content.AddShape(ShapeGeometryType.RegularPentagon, 170, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.ForestGreen));
        slide.Content.AddShape(ShapeGeometryType.Hexagon, 310, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.GreenYellow));
        slide.Content.AddShape(ShapeGeometryType.Octagon, 450, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.LightSeaGreen));

        slide.Content.AddShape(ShapeGeometryType.UpArrow, 30, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed));
        slide.Content.AddShape(ShapeGeometryType.RightArrow, 170, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.IndianRed));
        slide.Content.AddShape(ShapeGeometryType.UpDownArrow, 310, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.OrangeRed));
        slide.Content.AddShape(ShapeGeometryType.LeftRightArrow, 450, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.MediumVioletRed));

        presentation.Save("Shapes.pptx");
    }
}
