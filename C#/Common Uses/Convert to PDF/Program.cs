using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Example1();
        Example2();
    }

    static void Example1()
    {
        var presentation = PresentationDocument.Load("Reading.pptx");

        // In order to achieve the conversion of a loaded PowerPoint file to PDF,
        // we just need to save a PresentationDocument object to desired 
        // output file format.
        presentation.Save("Convert.pdf");
    }

    static void Example2()
    {
        PdfConformanceLevel conformanceLevel = PdfConformanceLevel.PdfA1a;

        // Load PowerPoint file.
        var presentation = PresentationDocument.Load("Reading.pptx");

        // Create PDF save options.
        var options = new PdfSaveOptions()
        {
            ConformanceLevel = conformanceLevel
        };

        // Save to PDF file.
        presentation.Save("Output.pdf", options);
    }
}