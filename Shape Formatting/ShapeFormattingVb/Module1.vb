Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create New slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create New "rounded rectangle" shape.
        Dim shape As Shape = slide.Content.AddShape(
            ShapeGeometryType.RoundedRectangle, 2, 2, 5, 4, LengthUnit.Centimeter)

        ' Get shape format.
        Dim format As ShapeFormat = shape.Format

        ' Get shape fill format.
        Dim fillFormat As FillFormat = format.Fill

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
        Dim lineFormat As LineFormat = shape.Format.Outline

        ' Set shape outline format as single solid red line.
        lineFormat.Fill.SetSolid(Color.FromName(ColorName.Red))
        lineFormat.DashType = LineDashType.Solid
        lineFormat.Width = Length.From(0.8, LengthUnit.Centimeter)
        lineFormat.CompoundType = LineCompoundType.Single

        presentation.Save("Shape Formatting.pdf")

    End Sub

End Module