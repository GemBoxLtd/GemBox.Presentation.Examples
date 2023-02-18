using GemBox.Presentation;

namespace PresentationMaui
{
    public partial class MainPage : ContentPage
    {
        static MainPage()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public MainPage()
        {
            InitializeComponent();
        }

        private string CreatePresentation()
        {
            var presentation = new PresentationDocument();

            presentation.Slides.AddNew(SlideLayoutType.Custom)
                .Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 10, 4, LengthUnit.Centimeter)
                .AddParagraph()
                .AddRun(text.Text);

            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Example.pptx");

            presentation.Save(filePath);

            return filePath;
        }

        private async void Button_Clicked(object sender, EventArgs e)
        {
            button.IsEnabled = false;
            activity.IsRunning = true;

            // In real apps the call to the method should be async (Task.Run(() => ....)
            var filePath = CreatePresentation();
            await Launcher.OpenAsync(new OpenFileRequest(Path.GetFileName(filePath), new ReadOnlyFile(filePath)));

            activity.IsRunning = false;
            button.IsEnabled = true;
        }
    }
}