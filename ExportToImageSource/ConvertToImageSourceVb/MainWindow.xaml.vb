Imports GemBox.Presentation

Class MainWindow

    Public Sub New()
        InitializeComponent()

        SetDocumentViewer(Me.ImageControl)
    End Sub

    Private Shared Sub SetDocumentViewer(image As Image)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument()

        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox As TextBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter)
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        Dim run As TextRun = textBox.AddParagraph().AddRun("Hello World!")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black))

        image.Source = presentation.ConvertToImageSource(SaveOptions.Image)

    End Sub

End Class
