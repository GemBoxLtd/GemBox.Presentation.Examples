using System;
using System.Linq;
using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = new PresentationDocument();

        // Create new master slide.
        MasterSlide master = presentation.MasterSlides.AddNew();

        // Create new empty layout slide for existing master slide.
        LayoutSlide layout = master.LayoutSlides.AddNew(SlideLayoutType.Custom);

        // Create title and subtitle placeholders on layout slide.
        layout.Content.AddPlaceholder(PlaceholderType.Title);
        layout.Content.AddPlaceholder(PlaceholderType.Subtitle);

        // Create new slide; will inherit all placeholders (title and subtitle) from template layout slide.
        Slide slide = presentation.Slides.AddNew(layout);

        // Retrieve "Title" placeholder.
        Shape shape = slide.Content.Drawings.OfType<Shape>().Where(item => item.Placeholder?.PlaceholderType == PlaceholderType.Title).First();

        // Set shape fill and outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));
        shape.Format.Fill.SetSolid(Color.FromName(ColorName.LightGray));

        // Set shape text.
        shape.Text.AddParagraph().AddRun("Placeholders, GemBox.Presentation");

        // Retrieve "Subtitle" placeholder.
        shape = slide.Content.Drawings.OfType<Shape>().Where(item => item.Placeholder?.PlaceholderType == PlaceholderType.Subtitle).First();

        // Set shape outline format.
        shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray));

        // Set shape text.
        shape.Text.AddParagraph().AddRun("Shows how to use placeholders with GemBox.Presentation component.");

        presentation.Save("Placeholders.pptx");
    }
}
