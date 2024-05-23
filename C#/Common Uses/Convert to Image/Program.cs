using GemBox.Presentation;
using System.IO;
using System.IO.Compression;

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

        // Load a PowerPoint file into the PresentationDocument object.
        var presentation = PresentationDocument.Load("Input.pptx");
        
        // Create image save options.
        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png)
        {
            SlideNumber = 0, // Select the first slide.
            Width = 1240 // Set the image width and keep the aspect ratio.
        };

        // Save the PresentationDocument object to a PNG file.
        presentation.Save("Output.png", imageOptions);
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load a PowerPoint file.
        var presentation = PresentationDocument.Load("Input.pptx");

        // Max integer value indicates that all presentation slides should be saved.
        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Tiff)
        {
            SlideCount = int.MaxValue
        };

        // Save the TIFF file with multiple frames, each frame represents a single PowerPoint slide.
        presentation.Save("Output.tiff", imageOptions);
    }

    static void Example3()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Load a PowerPoint file.
        var presentation = PresentationDocument.Load("Input.pptx");

        var imageOptions = new ImageSaveOptions();

        // Get PowerPoint pages, one for each slide.
        var pages = presentation.GetPaginator().Pages;

        // Create a ZIP file for storing PNG files.
        using (var archiveStream = File.OpenWrite("Output.zip"))
        using (var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create))
        {
            // Iterate through the PowerPoint pages.
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                PresentationDocumentPage page = pages[pageIndex];

                // Create a ZIP entry for each slide.
                var entry = archive.CreateEntry($"Slide {pageIndex + 1}.png");

                // Save each slide as a PNG image to the ZIP entry.
                using (var imageStream = new MemoryStream())
                using (var entryStream = entry.Open())
                {
                    page.Save(imageStream, imageOptions);
                    imageStream.CopyTo(entryStream);
                }
            }
        }
    }
}
