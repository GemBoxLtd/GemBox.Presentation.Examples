Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim password As String = "pass"
        Dim ownerPassword As String = ""

        Dim options = New PdfSaveOptions() With
        {
            .DocumentOpenPassword = password,
            .PermissionsPassword = ownerPassword,
            .Permissions = PdfPermissions.None
        }

        presentation.Save("PDF Encryption.pdf", options)
    End Sub
End Module
