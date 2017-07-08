using System;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = new PresentationDocument();

        // Get slide size.
        SlideSize size = presentation.SlideSize;

        // Set slide size.
        size.SizedFor = SlideSizeType.OnscreenShow16X10;
        size.Orientation = Orientation.Landscape;
        size.NumberSlidesFrom = 1;

        // Create new master slide.
        MasterSlide master = presentation.MasterSlides.AddNew();

        // Create new layout slide for existing master slide.
        LayoutSlide layout = master.LayoutSlides.AddNew(SlideLayoutType.TitleAndObject);

        // Create new slide from existing template layout slide.
        Slide slide = presentation.Slides.AddNew(layout);

        // If master slide collection is empty, this method will add a new master slide.
        // If layout slide collection of the last master slide doesn't contain a layout slide with the specified type, 
        // then a new layout slide with the specified type will be added.
        slide = presentation.Slides.AddNew(SlideLayoutType.TwoTextAndTwoObjects);

        presentation.Save("Slides.pptx");
    }
}
