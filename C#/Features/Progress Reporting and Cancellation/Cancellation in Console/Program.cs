using GemBox.Presentation;
using System;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // Create presentation.
        var presentation = new PresentationDocument();
        for (var i = 0; i < 10000; i++)
        {
            var slide = presentation.Slides.AddNew();
            var textBox = slide.Content.AddTextBox(100, 100, 100, 100);
            textBox.AddParagraph().AddRun(i.ToString());
        }

        var stopwatch = new Stopwatch();
        stopwatch.Start();

        // Create save options.
        var saveOptions = new PptxSaveOptions();
        saveOptions.ProgressChanged += (sender, args) =>
        {
            // Cancel operation after five seconds.
            if (stopwatch.Elapsed.Seconds >= 5)
                args.CancelOperation();
        };

        try
        {
            presentation.Save("Cancellation.pptx", saveOptions);
            Console.WriteLine("Operation fully finished");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Operation was cancelled");
        }
    }
}
