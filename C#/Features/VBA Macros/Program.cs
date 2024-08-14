using GemBox.Presentation;
using GemBox.Presentation.Vba;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();
        presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create the module.
        VbaModule vbaModule = presentation.VbaProject.Modules.Add("SampleModule");
        vbaModule.Code =
@"Sub SayHello()
    MsgBox ""Hello World!""
End Sub";

        // Save the presentation as macro-enabled PowerPoint file.
        presentation.Save("AddVbaModule.pptm");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("SampleVba.pptm");

        // Get the module.
        VbaModule vbaModule = presentation.VbaProject.Modules["Slide1"];
        // Update text for the popup message.
        vbaModule.Code = vbaModule.Code.Replace("Hello world!", "Hello from GemBox.Presentation!");

        presentation.Save("UpdateVbaModule.pptm");
    }
}
