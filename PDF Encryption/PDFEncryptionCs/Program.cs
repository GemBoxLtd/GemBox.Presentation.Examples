using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        string password = "pass";
        string ownerPassword = "";

        var options = new PdfSaveOptions()
        {
            DocumentOpenPassword = password,
            PermissionsPassword = ownerPassword,
            Permissions = PdfPermissions.None
        };

        presentation.Save("PDF Encryption.pdf", options);
    }
}