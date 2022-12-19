using System;
using System.Linq;
using System.Text;
using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        var sb = new StringBuilder();

        var slide = presentation.Slides[0];

        foreach (var shape in slide.Content.Drawings.OfType<Shape>())
        {
            sb.AppendFormat("Shape ShapeType={0}:", shape.ShapeType);
            sb.AppendLine();

            foreach (var paragraph in shape.Text.Paragraphs)
            {
                foreach (var run in paragraph.Elements.OfType<TextRun>())
                {
                    var isBold = run.Format.Bold;
                    var text = run.Text;

                    sb.AppendFormat("{0}{1}{2}", isBold ? "<b>" : "", text, isBold ? "</b>" : "");
                }

                sb.AppendLine();
            }

            sb.AppendLine("----------");
        }

        Console.WriteLine(sb.ToString());
    }
}