using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using GemBox.Presentation;

public partial class MainWindow : Window
{
    XpsDocument xpsDocument;

    public MainWindow()
    {
        InitializeComponent();

        this.SetDocumentViewer(this.DocumentViewer);
    }

    private void SetDocumentViewer(DocumentViewer documentViewer)
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter);
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        var run = textBox.AddParagraph().AddRun("Hello World!");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black));

        // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        this.xpsDocument = presentation.ConvertToXpsDocument(SaveOptions.Xps);

        documentViewer.Document = this.xpsDocument.GetFixedDocumentSequence();
    }
}