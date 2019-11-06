using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("EmbeddedObjects.pptx");

        presentation.Save("Embedded Objects.pptx");
    }
}