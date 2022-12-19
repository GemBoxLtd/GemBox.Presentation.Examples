Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim inputPassword As String = "inpass"
        Dim outputPassword As String = "outpass"

        Dim presentation = PresentationDocument.Load("PptxEncryption.pptx", New PptxLoadOptions With {.Password = inputPassword})
        presentation.Save("PPTX Encryption.pptx", New PptxSaveOptions With {.Password = outputPassword})
    End Sub
End Module