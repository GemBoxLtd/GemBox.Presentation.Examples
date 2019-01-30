Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("EmbeddedObjects.pptx")

        presentation.Save("Embedded Objects.pptx")
    End Sub
End Module