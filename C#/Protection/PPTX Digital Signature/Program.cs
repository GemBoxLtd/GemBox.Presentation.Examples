using GemBox.Presentation;
using GemBox.Presentation.Security;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
        Example3();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        var saveOptions = new PptxSaveOptions();
        saveOptions.DigitalSignatures.Add(new PptxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxECDsa521.pfx",
            CertificatePassword = "GemBoxPassword"
        });

        presentation.Save("PPTX Digital Signature.pptx", saveOptions);
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        var signature1 = new PptxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxECDsa521.pfx",
            CertificatePassword = "GemBoxPassword",
            CommitmentType = DigitalSignatureCommitmentType.Created,
            SignerRole = "Developer"
        };
        // Embed intermediate certificate.
        signature1.Certificates.Add(new Certificate("GemBoxECDsa.crt"));

        var signature2 = new PptxDigitalSignatureSaveOptions()
        {
            CertificatePath = "GemBoxRSA4096.pfx",
            CertificatePassword = "GemBoxPassword",
            CommitmentType = DigitalSignatureCommitmentType.Approved,
            SignerRole = "Manager"
        };
        // Embed intermediate certificate.
        signature2.Certificates.Add(new Certificate("GemBoxRSA.crt"));

        var saveOptions = new PptxSaveOptions();
        saveOptions.DigitalSignatures.Add(signature1);
        saveOptions.DigitalSignatures.Add(signature2);

        presentation.Save("PPTX Digital Signatures.pptx", saveOptions);
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load signed presentation file.
        var presentation = PresentationDocument.Load("Signed.pptx");

        // Signature is removed by simply saving the presentation with default save options.
        presentation.Save("Unsigned.pptx");
    }
}
