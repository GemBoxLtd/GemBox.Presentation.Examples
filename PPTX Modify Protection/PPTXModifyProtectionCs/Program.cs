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

        slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 20, 2, LengthUnit.Centimeter)
            .AddParagraph()
            .AddRun("This presentation has been opened in read-only mode, no changes can be made to a slide.");

        // ModifyProtection class is supported only for PPTX file format.
        ModifyProtection protection = presentation.ModifyProtection;
        protection.SetPassword("1234");

        presentation.Save("PPTX Modify Protection.pptx");
    }
}
