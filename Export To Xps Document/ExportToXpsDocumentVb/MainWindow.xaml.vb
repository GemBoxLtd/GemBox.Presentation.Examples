Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Xps.Packaging
Imports GemBox.Presentation

Class MainWindow

    Dim xpsDoc As XpsDocument

    Public Sub New()

        InitializeComponent()

        Me.SetDocumentViewer(Me.DocumentViewer)
    End Sub

    Private Sub SetDocumentViewer(documentViewer As DocumentViewer)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = New PresentationDocument()

        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 8, 4, LengthUnit.Centimeter)
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        Dim run = textBox.AddParagraph().AddRun("Hello World!")
        run.Format.Fill.SetSolid(Color.FromName(ColorName.Black))

        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        Me.xpsDoc = presentation.ConvertToXpsDocument(SaveOptions.Xps)

        documentViewer.Document = Me.xpsDoc.GetFixedDocumentSequence()
    End Sub
End Class