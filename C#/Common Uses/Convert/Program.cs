using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        // In order to achieve the conversion of a loaded PowerPoint file to PDF,
        // we just need to save a PresentationDocument object to desired 
        // output file format.
        presentation.Save("Convert.pdf");
    }
}