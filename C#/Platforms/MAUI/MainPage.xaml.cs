using GemBox.Presentation;
using System;
using System.IO;
using System.Threading.Tasks;

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

        private async Task<string> CreatePresentationAsync()
        {
            var presentation = new PresentationDocument();

            presentation.Slides.AddNew(SlideLayoutType.Custom)
                .Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 10, 4, LengthUnit.Centimeter)
                .AddParagraph()
                .AddRun(text.Text);

            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Example.pptx");

            await Task.Run(() => presentation.Save(filePath));

            return filePath;
        }

        private async void Button_Clicked(object sender, EventArgs e)
        {
            button.IsEnabled = false;
            activity.IsRunning = true;

            try
            {
                var filePath = await CreatePresentationAsync();
                await Launcher.OpenAsync(new OpenFileRequest(Path.GetFileName(filePath), new ReadOnlyFile(filePath)));
            }
            catch (Exception ex)
            {
                await DisplayAlert("Error", ex.Message, "Close");
            }

            activity.IsRunning = false;
            button.IsEnabled = true;
        }
    }
}