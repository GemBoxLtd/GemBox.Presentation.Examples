using System;
using System.Linq;
using System.Text;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = PresentationDocument.Load("Reading.pptx");

        StringBuilder sb = new StringBuilder();

        Slide slide = presentation.Slides[0];

        foreach (Shape shape in slide.Content.Drawings.OfType<Shape>())
        {
            sb.AppendFormat("Shape ShapeType={0}:", shape.ShapeType);
            sb.AppendLine();

            foreach (TextParagraph paragraph in shape.Text.Paragraphs)
            {
                foreach (TextRun run in paragraph.Elements.OfType<TextRun>())
                {
                    bool isBold = run.Format.Bold;
                    string text = run.Text;

                    sb.AppendFormat("{0}{1}{2}", isBold ? "<b>" : "", text, isBold ? "</b>" : "");
                }

                sb.AppendLine();
            }

            sb.AppendLine("----------");
        }

        Console.WriteLine(sb.ToString());
    }
}
