using System.Windows;
using System.Windows.Controls;
using GemBox.Presentation;

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

        var presentation = new PresentationDocument();

        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        var textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter);
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        var run = textBox.AddParagraph().AddRun("Hello World!");
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black));

        image.Source = presentation.ConvertToImageSource(SaveOptions.Image);
    }
}