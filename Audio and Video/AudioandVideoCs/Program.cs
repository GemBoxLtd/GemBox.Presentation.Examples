using System;
using System.IO;
using GemBox.Presentation;
using GemBox.Presentation.Media;

class Program
{
    static void Main(string[] args)
    {
        // If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        PresentationDocument presentation = new PresentationDocument();

        string pathToResources = "Resources";

        // Create new presentation slide.
        Slide slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create and add audio content.
        AudioContent audio = null;
        using (var stream = File.OpenRead(Path.Combine(pathToResources, "Applause.wav")))
            audio = slide.Content.AddAudio(AudioContentType.Wav, stream, 2, 2, LengthUnit.Centimeter);

        // Set the ending fade durations for the media.
        audio.Fade.End = TimeOffset.From(300, TimeOffsetUnit.Millisecond);

        // Get the picture associated with this media.
        Picture picture = audio.Picture;

        // Set drawing properties.
        picture.Action.Click.Set(ActionType.PlayMedia);
        picture.Layout.Width = Length.From(7, LengthUnit.Centimeter);
        picture.Layout.Height = Length.From(7, LengthUnit.Centimeter);
        picture.Name = "Applause.wav";

        // Create and add video content.
        VideoContent video = null;
        using (var stream = File.OpenRead(Path.Combine(pathToResources, "Wildlife.wmv")))
            video = slide.Content.AddVideo("video/x-ms-wmv", stream, 10, 2, 10, 5.6, LengthUnit.Centimeter);

        // Set drawing properties.
        video.Picture.Action.Click.Set(ActionType.PlayMedia);
        video.Picture.Name = "Wildlife.wmv";

        // Set the amount of time to be trimmed from the start and end of the media.
        video.Trim.Start = TimeOffset.From(600, TimeOffsetUnit.Millisecond);
        video.Trim.End = TimeOffset.From(800, TimeOffsetUnit.Millisecond);

        // Set the starting and ending fade durations for the media.
        video.Fade.Start = TimeOffset.From(100, TimeOffsetUnit.Millisecond);
        video.Fade.End = TimeOffset.From(200, TimeOffsetUnit.Millisecond);

        // Add video bookmarks.
        video.Bookmarks.Add(TimeOffset.From(1500, TimeOffsetUnit.Millisecond));
        video.Bookmarks.Add(TimeOffset.From(3000, TimeOffsetUnit.Millisecond));

        presentation.Save("Audio and Video.pptx");
    }
}
