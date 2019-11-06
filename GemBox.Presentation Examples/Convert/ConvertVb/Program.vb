Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        ' In order to achieve the conversion of a loaded PowerPoint file to PDF,
        ' we just need to save a PresentationDocument object to desired 
        ' output file format.
        presentation.Save("Convert.pdf")
    End Sub
End Module