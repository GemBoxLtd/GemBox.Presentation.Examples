Imports System.Windows.Xps.Packaging
Imports GemBox.Presentation

Class MainWindow

    Dim xpsDocument As XpsDocument

    Public Sub New()
        InitializeComponent()

        SetDocumentViewer(Me.DocumentViewer)
    End Sub

    Private Shared Sub SetDocumentViewer(documentViewer As DocumentViewer)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument()

        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox As TextBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter)
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        Dim run As TextRun = textBox.AddParagraph().AddRun("Hello World!")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black))

        Dim xpsDocument = presentation.ConvertToXpsDocument(SaveOptions.Xps)
        documentViewer.Document = xpsDocument.GetFixedDocumentSequence()

        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        xpsDocument = presentation.ConvertToXpsDocument(SaveOptions.Xps)

        documentViewer.Document = xpsDocument.GetFixedDocumentSequence()

    End Sub

End Class
