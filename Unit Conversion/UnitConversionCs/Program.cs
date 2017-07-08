using System;
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

        sb.AppendLine("Slide size (width X height):");

        Length width = presentation.SlideSize.Width;
        Length height = presentation.SlideSize.Height;

        foreach (LengthUnit unit in Enum.GetValues(typeof(LengthUnit)))
        {
            sb.AppendFormat(
                "{0} X {1} {2}",
                width.To(unit),
                height.To(unit),
                unit.ToString().ToLowerInvariant());

            sb.AppendLine();
        }

        Console.WriteLine(sb.ToString());
    }
}
