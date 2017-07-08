Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)
        Dim shape As Shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 2, 2, 8, 4, LengthUnit.Centimeter)
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.DarkBlue))

        Dim run As TextRun = shape.Text.AddParagraph().AddRun("This sample shows how to write or save a new PowerPoint file with GemBox.Presentation.")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.White))

        presentation.Save("Writing.pptx")

    End Sub

End Module