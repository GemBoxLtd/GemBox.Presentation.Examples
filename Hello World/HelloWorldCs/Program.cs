using System;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        
        PresentationDocument presentation = new PresentationDocument();

        Slide slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        TextBox textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter);

        TextParagraph paragraph = textBox.AddParagraph();

        paragraph.AddRun("Hello World!");

        presentation.Save("Hello World.pptx");
    }
}
