using System.Windows;
using System.Windows.Controls;
using GemBox.Presentation;

namespace ConvertToImageSourceCs
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            SetImageSource(this.ImageControl);
        }

        private static void SetImageSource(Image image)
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            PresentationDocument presentation = new PresentationDocument();

            Slide slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

            GemBox.Presentation.TextBox textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter);
            textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

            TextRun run = textBox.AddParagraph().AddRun("Hello World!");
            run.Format.Fill.SetSolid(Color.FromName(ColorName.Black));

            image.Source = presentation.ConvertToImageSource(SaveOptions.Image);
        }
    }
}
