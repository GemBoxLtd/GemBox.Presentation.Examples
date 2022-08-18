Imports System.IO
Imports GemBox.Pdf.Forms
Imports GemBox.Pdf.Security
Imports GemBox.Presentation

Module Program

    Sub Main()

        PAdES_B_B()

        PAdES_B_LTA()
    End Sub

    Sub PAdES_B_B()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        ' Create visual representation of digital signature on the first slide.
        Dim signature As Picture = Nothing
        Using stream = File.OpenRead("GemBoxSignature.png")
            signature = presentation.Slides(0).Content.AddPicture(
                PictureContentType.Png, stream, 25, 15, 4, 1, LengthUnit.Centimeter)
        End Using

        Dim options As New PdfSaveOptions() With
        {
            .DigitalSignature = New PdfDigitalSignatureSaveOptions() With
            {
                .CertificatePath = "GemBoxECDsa521.pfx",
                .CertificatePassword = "GemBoxPassword",
                .Signature = signature,
                .IsAdvancedElectronicSignature = True
            }
        }

        presentation.Save("PDF Digital Signature.pdf", options)
    End Sub

    Sub PAdES_B_LTA()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        ' Create visual representation of digital signature on the first slide.
        Dim signature As Picture = Nothing
        Using stream = File.OpenRead("GemBoxSignature.png")
            signature = presentation.Slides(0).Content.AddPicture(
                PictureContentType.Png, stream, 25, 15, 4, 1, LengthUnit.Centimeter)
        End Using

        ' If using Professional version, put your serial key below.
        GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Get a digital ID from PKCS#12/PFX file.
        Dim digitalId = New PdfDigitalId("GemBoxECDsa521.pfx", "GemBoxPassword")

        ' Create a PDF signer that will create PAdES B-LTA level signature.
        Dim signer = New PdfSigner(digitalId)

        ' PdfSigner should create CAdES-equivalent signature.
        signer.SignatureFormat = PdfSignatureFormat.CAdES

        ' PdfSigner will embed a timestamp created by freeTSA.org Time Stamp Authority in the signature.
        signer.Timestamper = New PdfTimestamper("https://freetsa.org/tsr")

        ' Make sure that all properties specified on PdfSigner are according to PAdES B-LTA level.
        signer.SignatureLevel = PdfSignatureLevel.PAdES_B_LTA

        ' Inject PdfSigner from GemBox.Pdf into
        ' PdfDigitalSignatureSaveOptions from GemBox.Presentation.
        Dim signatureOptions = PdfDigitalSignatureSaveOptions.FromSigner(
            Function() signer.SignatureFormat.ToString(),
            Function() signer.EstimatedSignatureContentsLength,
            Function(pdfFileStream) signer.ComputeSignature(pdfFileStream))

        signatureOptions.Signature = signature

        Dim options = New PdfSaveOptions() With
        {
            .DigitalSignature = signatureOptions
        }

        presentation.Save("PAdES B-LTA.pdf", options)

        Using pdfDocument = GemBox.Pdf.PdfDocument.Load("PAdES B-LTA.pdf")

            Dim signatureField = CType(pdfDocument.Form.Fields(0), PdfSignatureField)

            ' Download validation-related information for the signature and the signature's timestamp and embed it in the PDF file.
            ' This will make the signature "LTV enabled".
            pdfDocument.SecurityStore.AddValidationInfo(signatureField.Value)

            ' Add an invisible signature field to the PDF document that will hold the document timestamp.
            Dim timestampField = pdfDocument.Form.Fields.AddSignature()

            ' Initiate timestamping of a PDF file with the specified timestamper.
            timestampField.Timestamp(signer.Timestamper)

            ' Save any changes done to the PDF file that were done since the last time Save was called and
            ' finish timestamping of a PDF file.
            pdfDocument.Save()
        End Using
    End Sub
End Module