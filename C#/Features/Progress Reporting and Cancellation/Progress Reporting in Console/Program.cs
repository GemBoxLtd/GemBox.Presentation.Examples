using GemBox.Presentation;
using System;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        Console.WriteLine("Creating presentation");

        // Create large presentation.
        var presentation = new PresentationDocument();
        for (var i = 0; i < 10000; i++)
        {
            var slide = presentation.Slides.AddNew();
            var textBox = slide.Content.AddTextBox(100, 100, 100, 100);
            textBox.AddParagraph().AddRun(i.ToString());
        }

        // Create save options.
        var saveOptions = new PptxSaveOptions();
        saveOptions.ProgressChanged += (eventSender, args) =>
        {
            Console.WriteLine($"Progress changed - {args.ProgressPercentage}%");
        };

        // Save document.
        presentation.Save("presentation.pptx", saveOptions);
    }
}
