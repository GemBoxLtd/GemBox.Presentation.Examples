Imports GemBox.Presentation
Imports GemBox.Presentation.Security

Module Program

    Sub Main()
        Example1()
        Example2()
        Example3()
    End Sub

    Sub Example1()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

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
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

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

    Sub Example3()
        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        ' Load signed presentation file.
        Dim presentation = PresentationDocument.Load("Signed.pptx")

        ' Signature is removed by simply saving the presentation with default save options.
        presentation.Save("Unsigned.pptx")
    End Sub

End Module
