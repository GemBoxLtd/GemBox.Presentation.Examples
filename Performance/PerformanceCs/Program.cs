using System;
using System.Diagnostics;
using System.Linq;
using GemBox.Presentation;

namespace PerformanceCs
{
    class Program
    {
        static void Main(string[] args)
        {
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // If sample exceeds Free version limitations then continue as trial version: 
            // https://www.gemboxsoftware.com/presentation/help/html/Evaluation_and_Licensing.htm
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

            Console.WriteLine("Performance sample:");
            Console.WriteLine();

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            PresentationDocument presentation = PresentationDocument.Load("Template.pptx", LoadOptions.Pptx);

            Console.WriteLine("Load file (seconds): " + stopwatch.Elapsed.TotalSeconds);

            stopwatch.Reset();
            stopwatch.Start();

            int numberOfShapes = 0;
            int numberOfParagraphs = 0;

            foreach (Slide slide in presentation.Slides)
                foreach (Shape shape in slide.Content.Drawings.OfType<Shape>())
                {
                    foreach (TextParagraph paragraph in shape.Text.Paragraphs)
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
}
