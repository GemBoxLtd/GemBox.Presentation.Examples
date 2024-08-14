using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using GemBox.Presentation;
using System.Collections.Generic;
using System.IO;

[SimpleJob(RuntimeMoniker.Net80)]
[SimpleJob(RuntimeMoniker.Net48)]
public class Program
{
    private PresentationDocument presentation;
    private readonly Consumer consumer = new Consumer();

    public static void Main()
    {
        BenchmarkRunner.Run<Program>();
    }

    [GlobalSetup]
    public void SetLicense()
    {
        // If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");

        // If using Free version and example exceeds its limitations, use Trial or Time Limited version:
        // https://www.gemboxsoftware.com/presentation/examples/free-trial-professional/901

        this.presentation = PresentationDocument.Load("RandomSlides.pptx");
    }

    [Benchmark]
    public PresentationDocument Reading()
    {
        return PresentationDocument.Load("RandomSlides.pptx");
    }

    [Benchmark]
    public void Writing()
    {
        using (var stream = new MemoryStream())
            this.presentation.Save(stream, new PptxSaveOptions());
    }

    [Benchmark]
    public void Iterating()
    {
        this.LoopThroughAllDrawings().Consume(this.consumer);
    }

    public IEnumerable<Drawing> LoopThroughAllDrawings()
    {
        foreach (var slide in this.presentation.Slides)
            foreach (var drawing in slide.Content.Drawings.All())
                yield return drawing;
    }
}
