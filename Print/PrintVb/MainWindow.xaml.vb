Imports GemBox.Presentation
Imports Microsoft.Win32

Class MainWindow

    Dim presentation As PresentationDocument

    Public Sub New()
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        InitializeComponent()

        Me.EnableControls()
    End Sub

    Private Sub LoadFileBtn_Click(sender As Object, e As RoutedEventArgs)

        Dim fileDialog = New OpenFileDialog()
        fileDialog.Filter = "PPTX files (*.pptx, *.pptm, *.potx, *.potm)|*.pptx;*.pptm;*.potx;*.potm"

        If (fileDialog.ShowDialog() = True) Then
            Me.presentation = PresentationDocument.Load(fileDialog.FileName)

            Me.ShowPrintPreview()
            Me.EnableControls()
        End If

    End Sub

    Private Sub SimplePrint_Click(sender As Object, e As RoutedEventArgs)

        ' Print to default printer using default options
        Me.presentation.Print()

    End Sub

    Private Sub AdvancedPrint_Click(sender As Object, e As RoutedEventArgs)

        ' We can use PrintDialog for defining print options
        Dim printDialog = New PrintDialog()
        printDialog.UserPageRangeEnabled = True

        If (printDialog.ShowDialog() = True) Then

            Dim printOptions = New PrintOptions(printDialog.PrintTicket.GetXmlStream())

            printOptions.FromSlide = printDialog.PageRange.PageFrom - 1
            If (printDialog.PageRange.PageTo = 0) Then
                printOptions.ToSlide = Int32.MaxValue
            Else
                printOptions.ToSlide = printDialog.PageRange.PageTo - 1
            End If

            Me.presentation.Print(printDialog.PrintQueue.FullName, printOptions)
        End If

    End Sub

    ' We can use DocumentViewer for print preview (but we don't need).
    Private Sub ShowPrintPreview()

        ' XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        ' Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        Dim xpsDocument = presentation.ConvertToXpsDocument(SaveOptions.Xps)
        Me.DocViewer.Tag = xpsDocument

        Me.DocViewer.Document = xpsDocument.GetFixedDocumentSequence()

    End Sub

    Private Sub EnableControls()

        Dim isEnabled As Boolean = Me.presentation IsNot Nothing

        Me.DocViewer.IsEnabled = isEnabled
        Me.SimplePrintFileBtn.IsEnabled = isEnabled
        Me.AdvancedPrintFileBtn.IsEnabled = isEnabled

    End Sub

End Class
