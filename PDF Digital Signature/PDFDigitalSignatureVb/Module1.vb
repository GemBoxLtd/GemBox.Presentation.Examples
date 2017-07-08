Imports System
Imports System.IO
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = PresentationDocument.Load("Reading.pptx")

        Dim pathToResources As String = "Resources"

        Dim slide As Slide = presentation.Slides(0)

        Dim signature As Picture = Nothing
        Using stream As Stream = File.OpenRead(
            Path.Combine(pathToResources, "GemBoxSignature.png"))
            signature = slide.Content.AddPicture(
                PictureContentType.Png, stream, 25, 15, 4, 1, LengthUnit.Centimeter)
        End Using

        Dim options = New PdfSaveOptions()
        Dim digitalSignature = options.DigitalSignature

        digitalSignature.CertificatePath = Path.Combine(
            pathToResources, "GemBoxSampleExplorer.pfx")
        digitalSignature.CertificatePassword = "GemBoxPassword"
        digitalSignature.Signature = signature

        presentation.Save("PDF Digital Signature.pdf", options)

    End Sub

End Module