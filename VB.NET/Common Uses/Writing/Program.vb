Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument

        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim shape = slide.Content.AddShape(ShapeGeometryType.RoundedRectangle, 2, 2, 8, 4, LengthUnit.Centimeter)
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.DarkBlue))

        Dim run = shape.Text.AddParagraph().AddRun("This sample shows how to write or save a new PowerPoint file with GemBox.Presentation.")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.White))

        presentation.Save("Writing.pptx")
    End Sub
End Module