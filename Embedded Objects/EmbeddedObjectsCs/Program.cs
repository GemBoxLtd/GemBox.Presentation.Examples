using System;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = PresentationDocument.Load("EmbeddedObjects.pptx");
        
        presentation.Save("Embedded Objects.pptx");
    }
}
