using GemBox.Presentation;
using System;
using System.Text;

class Program
{
    static void Main()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Reading.pptx");

        var sb = new StringBuilder();

        sb.AppendLine("Slide size (width X height):");

        var width = presentation.SlideSize.Width;
        var height = presentation.SlideSize.Height;

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
