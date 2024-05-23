Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        ' Create New slide.
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create New "rounded rectangle" shape.
        Dim shape = slide.Content.AddShape(
            ShapeGeometryType.RoundedRectangle, 2, 2, 5, 4, LengthUnit.Centimeter)

        ' Get shape format.
        Dim format = shape.Format

        ' Get shape fill format.
        Dim fillFormat = format.Fill

        ' Set shape fill format as solid fill.
        fillFormat.SetSolid(Color.FromName(ColorName.DarkBlue))

        ' Create new "rectangle" shape.
        shape = slide.Content.AddShape(
            ShapeGeometryType.Rectangle, 8, 2, 5, 4, LengthUnit.Centimeter)

        ' Set shape fill format as solid fill.
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.Yellow))

        ' Set shape outline format as solid fill.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.Green))

        ' Create new "rounded rectangle" shape.
        shape = slide.Content.AddShape(
            ShapeGeometryType.RoundedRectangle, 14, 2, 5, 4, LengthUnit.Centimeter)

        ' Set shape fill format as no fill.
        shape.Format.Fill.SetNone()

        ' Get shape outline format.
        Dim lineFormat = shape.Format.Outline

        ' Set shape outline format as single solid red line.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red))
        lineFormat.DashType = LineDashType.Solid
        lineFormat.Width = Length.From(0.8, LengthUnit.Centimeter)
        lineFormat.CompoundType = LineCompoundType.Single

        presentation.Save("Shape Formatting.pptx")
    End Sub
End Module
