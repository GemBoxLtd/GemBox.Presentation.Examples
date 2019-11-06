using System.Windows;
using System.Windows.Controls;
using System.Windows.Xps.Packaging;
using GemBox.Presentation;
using Microsoft.Win32;

public partial class MainWindow : Window
{
    private PresentationDocument presentation;

    public MainWindow()
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        InitializeComponent();

        this.EnableControls();
    }

    private void LoadFileBtn_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog fileDialog = new OpenFileDialog();
        fileDialog.Filter = "PPTX files (*.pptx, *.pptm, *.potx, *.potm)|*.pptx;*.pptm;*.potx;*.potm";

        if (fileDialog.ShowDialog() == true)
        {
            this.presentation = PresentationDocument.Load(fileDialog.FileName);

            this.ShowPrintPreview();
            this.EnableControls();
        }
    }

    private void SimplePrint_Click(object sender, RoutedEventArgs e)
    {
        // Print to default printer using default options
        this.presentation.Print();
    }

    private void AdvancedPrint_Click(object sender, RoutedEventArgs e)
    {
        // We can use PrintDialog for defining print options
        PrintDialog printDialog = new PrintDialog();
        printDialog.UserPageRangeEnabled = true;

        if (printDialog.ShowDialog() == true)
        {
            PrintOptions printOptions = new PrintOptions(printDialog.PrintTicket.GetXmlStream());

            printOptions.FromSlide = printDialog.PageRange.PageFrom - 1;
            printOptions.ToSlide = printDialog.PageRange.PageTo == 0 ? int.MaxValue : printDialog.PageRange.PageTo - 1;

            this.presentation.Print(printDialog.PrintQueue.FullName, printOptions);
        }
    }

    // We can use DocumentViewer for print preview (but we don't need).
    private void ShowPrintPreview()
    {
        // XpsDocument needs to stay referenced so that DocumentViewer can access additional required resources.
        // Otherwise, GC will collect/dispose XpsDocument and DocumentViewer will not work.
        XpsDocument xpsDocument = this.presentation.ConvertToXpsDocument(SaveOptions.Xps);
        this.DocViewer.Tag = xpsDocument;

        this.DocViewer.Document = xpsDocument.GetFixedDocumentSequence();
    }

    private void EnableControls()
    {
        var isEnabled = this.presentation != null;

        this.DocViewer.IsEnabled = isEnabled;
        this.SimplePrintFileBtn.IsEnabled = isEnabled;
        this.AdvancedPrintFileBtn.IsEnabled = isEnabled;
    }
}