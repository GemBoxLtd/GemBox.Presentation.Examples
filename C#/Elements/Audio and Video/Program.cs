using GemBox.Presentation;
using GemBox.Presentation.Media;
using System.IO;
using System.IO.Compression;
using System.Linq;

class Program
{
    static void Main()
    {
        Example1();
        Example2();
    }

    static void Example1()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = new PresentationDocument();

        // Create new presentation slide.
        var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

        // Create and add audio content.
        AudioContent audio = null;
        using (var stream = File.OpenRead("Applause.wav"))
            audio = slide.Content.AddAudio(AudioContentType.Wav, stream, 2, 2, LengthUnit.Centimeter);

        // Set the ending fade durations for the media.
        audio.Fade.End = TimeOffset.From(300, TimeOffsetUnit.Millisecond);

        // Get the picture associated with this media.
        var picture = audio.Picture;

        // Set drawing properties.
        picture.Action.Click.Set(ActionType.PlayMedia);
        picture.Layout.Width = Length.From(7, LengthUnit.Centimeter);
        picture.Layout.Height = Length.From(7, LengthUnit.Centimeter);
        picture.Name = "Applause.wav";

        // Create and add video content.
        VideoContent video = null;
        using (var stream = File.OpenRead("Wildlife.wmv"))
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

    static void Example2()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        var presentation = PresentationDocument.Load("Input Audio and Video.pptx");
        var slide = presentation.Slides[0];

        // Get audios from first slide.
        var audios = slide.Content.Drawings.All()
            .OfType<Picture>()
            .Where(p => p.Media?.MediaType == MediaType.Audio)
            .Select(p => p.Media as AudioContent);

        // Get videos from first slide.
        var videos = slide.Content.Drawings.All()
            .OfType<Picture>()
            .Where(p => p.Media?.MediaType == MediaType.Video)
            .Select(p => p.Media as VideoContent);

        // Create a ZIP file for storing audio and video files.
        using (var archiveStream = File.OpenWrite("Output.zip"))
        using (var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create))
        {
            int counter = 0;
            foreach (var audio in audios)
            {
                string extension = audio.Content.ContentType.Replace("audio/", string.Empty);
                var entry = archive.CreateEntry($"Audio {++counter}.{extension}");

                // Export audio from PowerPoint file to the ZIP entry.
                using (var entryStream = entry.Open())
                using (var audioStream = audio.Content.Open())
                    audioStream.CopyTo(entryStream);
            }

            counter = 0;
            foreach (var video in videos)
            {
                string extension = video.Content.ContentType.Replace("video/", string.Empty);
                var entry = archive.CreateEntry($"Video {++counter}.{extension}");

                // Export video from PowerPoint file to the ZIP entry.
                using (var entryStream = entry.Open())
                using (var videoStream = video.Content.Open())
                    videoStream.CopyTo(entryStream);
            }
        }
    }
}
