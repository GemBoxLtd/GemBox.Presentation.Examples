Imports System.Windows
Imports System.Windows.Controls
Imports GemBox.Presentation

Class MainWindow

    Public Sub New()

        InitializeComponent()

        SetImageSource(Me.ImageControl)
    End Sub

    Private Shared Sub SetImageSource(image As Image)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument()

        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter)
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        Dim run = textBox.AddParagraph().AddRun("Hello World!")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black))

        image.Source = presentation.ConvertToImageSource(SaveOptions.Image)
    End Sub
End Class