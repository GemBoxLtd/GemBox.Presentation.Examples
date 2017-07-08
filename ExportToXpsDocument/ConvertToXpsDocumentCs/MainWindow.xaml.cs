using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using GemBox.Presentation;

namespace ConvertToXpsDocumentCs
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
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

            PresentationDocument presentation = new PresentationDocument();

            Slide slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

            GemBox.Presentation.TextBox textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter);
            textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

            TextRun run = textBox.AddParagraph().AddRun("Hello World!");
            run.Format.Fill.SetSolid(Color.FromName(ColorName.Black));

            // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
            // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
            this.xpsDocument = presentation.ConvertToXpsDocument(SaveOptions.Xps);

            documentViewer.Document = this.xpsDocument.GetFixedDocumentSequence();
        }
    }
}
