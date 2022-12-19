Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("RightToLeft.pptx")

        Dim slide = presentation.Slides(0)

        Dim shape = slide.Content.AddShape(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter)

        ' Create a new right-to-left paragraph.
        Dim paragraph = shape.Text.AddParagraph()
        paragraph.Format.RightToLeft = true
        paragraph.Format.Alignment = HorizontalAlignment.Right
        Dim run = paragraph.AddRun("هذا ثمّة أمّا العالم، أم, السادس مواقعها")
        run.Format.Size = Length.From(28, LengthUnit.Point)

        presentation.Save("RightToLeft.pdf")
    End Sub

End Module