Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create New slide; will create "custom" layout slide And default master slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        slide.Content.AddShape(ShapeGeometryType.RectangularCallout, 30, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.AliceBlue))
        slide.Content.AddShape(ShapeGeometryType.RoundedRectangularCallout, 170, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.BlueViolet))
        slide.Content.AddShape(ShapeGeometryType.OvalCallout, 310, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.CadetBlue))
        slide.Content.AddShape(ShapeGeometryType.CloudCallout, 450, 30, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.CornflowerBlue))

        slide.Content.AddShape(ShapeGeometryType.ActionButtonEnd, 30, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.DarkSeaGreen))
        slide.Content.AddShape(ShapeGeometryType.ActionButtonForwardOrNext, 170, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.ForestGreen))
        slide.Content.AddShape(ShapeGeometryType.ActionButtonHelp, 310, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.GreenYellow))
        slide.Content.AddShape(ShapeGeometryType.ActionButtonHome, 450, 150, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.LightSeaGreen))

        slide.Content.AddShape(ShapeGeometryType.UpArrow, 30, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.DarkRed))
        slide.Content.AddShape(ShapeGeometryType.UpArrowCallout, 170, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.IndianRed))
        slide.Content.AddShape(ShapeGeometryType.UpDownArrow, 310, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.OrangeRed))
        slide.Content.AddShape(ShapeGeometryType.UpDownArrowCallout, 450, 270, 130, 100, LengthUnit.Point).Format.Fill.SetSolid(Color.FromName(ColorName.MediumVioletRed))

        presentation.Save("Shapes.pptx")

    End Sub

End Module