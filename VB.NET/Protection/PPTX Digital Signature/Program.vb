Imports GemBox.Presentation
Imports GemBox.Presentation.Security

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim saveOptions As New PptxSaveOptions()
        saveOptions.DigitalSignatures.Add(New PptxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxECDsa521.pfx",
            .CertificatePassword = "GemBoxPassword"
        })

        presentation.Save("PPTX Digital Signature.pptx", saveOptions)
    End Sub

    Sub Example2()
        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim signature1 As New PptxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxECDsa521.pfx",
            .CertificatePassword = "GemBoxPassword",
            .CommitmentType = DigitalSignatureCommitmentType.Created,
            .SignerRole = "Developer"
        }
        ' Embed intermediate certificate.
        signature1.Certificates.Add(New Certificate("GemBoxECDsa.crt"))

        Dim signature2 As New PptxDigitalSignatureSaveOptions() With
        {
            .CertificatePath = "GemBoxRSA4096.pfx",
            .CertificatePassword = "GemBoxPassword",
            .CommitmentType = DigitalSignatureCommitmentType.Approved,
            .SignerRole = "Manager"
        }
        ' Embed intermediate certificate.
        signature2.Certificates.Add(New Certificate("GemBoxRSA.crt"))

        Dim saveOptions As New PptxSaveOptions()
        saveOptions.DigitalSignatures.Add(signature1)
        saveOptions.DigitalSignatures.Add(signature2)

        presentation.Save("PPTX Digital Signatures.pptx", saveOptions)
    End Sub
End Module