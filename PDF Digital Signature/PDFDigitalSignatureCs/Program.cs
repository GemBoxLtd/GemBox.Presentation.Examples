using System;
using System.IO;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = PresentationDocument.Load("Reading.pptx");

        string pathToResources = "Resources";

        var slide = presentation.Slides[0];

        Picture signature = null;
        using (var stream = File.OpenRead(
            Path.Combine(pathToResources, "GemBoxSignature.png")))
            signature = slide.Content.AddPicture(
                PictureContentType.Png, stream, 25, 15, 4, 1, LengthUnit.Centimeter);

        var options = new PdfSaveOptions()
        {
            DigitalSignature =
            {
                CertificatePath = Path.Combine(
                    pathToResources, "GemBoxSampleExplorer.pfx"),
                CertificatePassword = "GemBoxPassword",
                Signature = signature
            }
        };

        presentation.Save("PDF Digital Signature.pdf", options);
    }
}
