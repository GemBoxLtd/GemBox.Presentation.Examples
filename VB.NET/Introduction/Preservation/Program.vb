Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load PowerPoint presentation; preservation feature is enabled by default.
        Dim presentation = PresentationDocument.Load("Preservation.pptx")
        Dim slide = presentation.Slides(0)

        ' Modify the presentation.
        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 80, 300, 300, 80, LengthUnit.Point)
        textBox.AddParagraph().AddRun("You can preserve unsupported features when modifying a presentation!")

        ' Save PowerPoint presentation to and output file of the same format together with
        ' preserved information (unsupported features) from the input file.
        presentation.Save("PreservedOutput.pptx")

    End Sub

End Module