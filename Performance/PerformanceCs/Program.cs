using System;
using System.Diagnostics;
using System.Linq;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If example exceeds Free version limitations then continue as trial version: 
        // https://www.gemboxsoftware.com/Presentation/help/html/Evaluation_and_Licensing.htm
        ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        Console.WriteLine("Performance example:");
        Console.WriteLine();

        var stopwatch = new Stopwatch();
        stopwatch.Start();

        var presentation = PresentationDocument.Load("Template.pptx", LoadOptions.Pptx);

        Console.WriteLine("Load file (seconds): " + stopwatch.Elapsed.TotalSeconds);

        stopwatch.Reset();
        stopwatch.Start();

        int numberOfShapes = 0;
        int numberOfParagraphs = 0;

        foreach (var slide in presentation.Slides)
            foreach (var shape in slide.Content.Drawings.OfType<Shape>())
            {
                foreach (var paragraph in shape.Text.Paragraphs)
                    ++numberOfParagraphs;

                ++numberOfShapes;
            }

        Console.WriteLine("Iterate through " + numberOfShapes + " shapes and " + numberOfParagraphs + " paragraphs (seconds): " + stopwatch.Elapsed.TotalSeconds);

        stopwatch.Reset();
        stopwatch.Start();

        presentation.Save("Report.pptx");

        Console.WriteLine("Save file (seconds): " + stopwatch.Elapsed.TotalSeconds);
    }
}