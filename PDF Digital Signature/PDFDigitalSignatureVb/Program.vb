Imports System.IO
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim slide = presentation.Slides(0)

        Dim signature As Picture = Nothing
        Using stream = File.OpenRead("GemBoxSignature.png")
            signature = slide.Content.AddPicture(
                PictureContentType.Png, stream, 25, 15, 4, 1, LengthUnit.Centimeter)
        End Using

        Dim options = New PdfSaveOptions()

        Dim digitalSignature = options.DigitalSignature

        digitalSignature.CertificatePath = "GemBoxExampleExplorer.pfx"
        digitalSignature.CertificatePassword = "GemBoxPassword"
        digitalSignature.Signature = signature

        presentation.Save("PDF Digital Signature.pdf", options)
    End Sub
End Module