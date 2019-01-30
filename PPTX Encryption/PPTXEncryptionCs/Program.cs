using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        string inputPassword = "inpass";
        string outputPassword = "outpass";

        var presentation = PresentationDocument.Load("PptxEncryption.pptx", new PptxLoadOptions() { Password = inputPassword });
        presentation.Save("PPTX Encryption.pptx", new PptxSaveOptions() { Password = outputPassword });
    }
}