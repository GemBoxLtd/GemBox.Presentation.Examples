using GemBox.Presentation;
using System.IO;

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

        // Load input HTML file.
        string html = File.ReadAllText("Input.html");

        // Create a presentation and add a new slide with text box.
        var presentation = new PresentationDocument();
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);
        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 0.5, 0.5, 32, 15, LengthUnit.Centimeter);

        // Loads text and styling from provided html into textBox.
        textBox.Shape.TextContent.LoadText(html, LoadOptions.Html);

        // Save the presentation to a PPTX file.
        presentation.Save("Load Text From Html.pptx");
    }

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // HTML content.
        string html =
@"<h2>Ordered list example</h2>
<ol style=""list-style-type: Decimal"">
    <li>First item</li>
    <li>Second item
        <ol style=""list-style-type: lower-alpha"">
            <li>Sub item 01
                <ol style=""list-style-type: lower-roman"">
                    <li>Sub item 01</li>
                    <li>Sub item 02</li>
                    <li>Sub item 03</li>
                </ol>
            </li>
            <li>Sub item 02</li>
            <li>Sub item 03</li>
        </ol>
    </li>
    <li>Third item</li>
    <li value=""50"">Arbitrary item</li>
    <li>Next item</li>
</ol>
<h2>Unordered list example</h2>
<ul style=""list-style-type: disc"">
    <li>First item</li>
    <li>Second item
        <ul style=""list-style-type: circle"">
            <li>Sub item 01
                <ul style=""list-style-type: square"">
                    <li>Sub item 01</li>
                    <li>Sub item 02</li>
                    <li>Sub item 02</li>
                </ul>
            </li>
            <li>Sub item 02</li>
            <li>Sub item 03</li>
        </ul>
    </li>
    <li>Third item</li>
</ul>";

        // Create a presentation and add a new slide with text box.
        var presentation = new PresentationDocument();
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);
        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 0.5, 0.5, 32, 15, LengthUnit.Centimeter);

        // Loads text and styling from provided html into textBox.
        textBox.Shape.TextContent.LoadText(html, LoadOptions.Html);

        // Save the presentation to a PPTX file.
        presentation.Save("Load Text From Html With List.pptx");
    }
}
