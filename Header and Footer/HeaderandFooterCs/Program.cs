using GemBox.Presentation;

class Program
{
    static void Main()
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new master slide.
        var master = presentation.MasterSlides.AddNew();
        master.Content.AddPlaceholder(PlaceholderType.Date);
        master.Content.AddPlaceholder(PlaceholderType.SlideNumber);

        // Set "DateTime" and "SlideNumber" placeholders visible on slides.
        master.HeaderFooter.IsDateTimeEnabled = true;
        master.HeaderFooter.IsSlideNumberEnabled = true;

        // Create new slides; will inherit "DateTime" and "SlideNumber" placeholders from master slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.VerticalTitleAndText);
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoObjects);
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoObjectsAndText);

        presentation.Save("Header and Footer.pptx");
    }
}