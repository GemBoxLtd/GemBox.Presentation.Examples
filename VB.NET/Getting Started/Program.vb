Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As New PresentationDocument()
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)
        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter)
        Dim paragraph = textBox.AddParagraph()

        paragraph.AddRun("Hello World!")

        presentation.Save("HelloWorld.pptx")

    End Sub

End Module