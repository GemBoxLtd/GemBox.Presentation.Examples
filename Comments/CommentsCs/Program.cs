using GemBox.Presentation;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new presentation slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Adds a new comment with a new author in the top-left corner of the slide.
        var comment = slide.Comments.Add("GBP", "GemBox.Presentation", "Shows how to use comments with GemBox.Presentation component.");

        // Change comment position.
        comment.Left = Length.From(50, LengthUnit.Centimeter);
        comment.Top = Length.From(10, LengthUnit.Centimeter);

        // Adds a new comment with the same author as the previously added comment.
        slide.Comments.Add("Another comment from GemBox.Presentation.");

        presentation.Save("Comments.pptx");
    }
}