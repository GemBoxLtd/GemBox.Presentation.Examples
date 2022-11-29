[![NuGet version](https://img.shields.io/nuget/v/GemBox.Presentation?style=for-the-badge)](https://www.nuget.org/packages/GemBox.Presentation/) [![NuGet downloads](https://img.shields.io/nuget/dt/GemBox.Presentation?style=for-the-badge)](https://www.nuget.org/packages/GemBox.Presentation/) [![Visual Studio Marketplace rating](https://img.shields.io/visual-studio-marketplace/stars/GemBoxSoftware.GemBoxPresentation?style=for-the-badge)](https://marketplace.visualstudio.com/items?itemName=GemBoxSoftware.GemBoxPresentation)

## What is GemBox.Presentation?

GemBox.Presentation is a .NET component that enables you to read, write, convert, and print presentation files (PPTX, PPT, and PPSX) from .NET applications.

With GemBox.Presentation you get a fast and reliable component that’s easy to use and doesn't depend on Microsoft PowerPoint. It requires only .NET and it’s much faster than Microsoft Office Automation!


## GemBox.Presentation Features

- [Read](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-open-read-powerpoint/201) PowerPoint (PPTX) and PowerPoint 97-2003 (PPT) presentations.
- [Write/create](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-create-write-powerpoint/202) and [convert](https://www.gemboxsoftware.com/presentation/examples/c-sharp-convert-powerpoint-to-pdf/204) to PPTX, PDF, PDF/A, XPS, and image formats.
- View presentations in [WPF](https://www.gemboxsoftware.com/presentation/examples/powerpoint-xpsdocument-wpf/1701) applications.
- [Print](https://www.gemboxsoftware.com/presentation/examples/c-sharp-print-powerpoint/251) presentations.
- [Encrypt PPTX](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-pptx-encryption/803) presentations. [Encrypt](https://www.gemboxsoftware.com/presentation/examples/pdf-encryption/801) and [digitally sign PDF](https://www.gemboxsoftware.com/presentation/examples/pdf-digital-signature/802) presentations.
- Get, create, or edit [master slides, layout slides, individual slides](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-powerpoint-slides/401), [master notes slide, and notes slides](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-powerpoint-slide-notes/411).
- Get, create, or edit drawings like [text boxes](https://www.gemboxsoftware.com/presentation/examples/powerpoint-textboxes/404), [connectors](https://www.gemboxsoftware.com/presentation/examples/powerpoint-shapes/403), [pictures](https://www.gemboxsoftware.com/presentation/examples/powerpoint-pictures/405), [tables](https://www.gemboxsoftware.com/presentation/examples/powerpoint-tables/601), [charts](https://www.gemboxsoftware.com/presentation/examples/powerpoint-charts/412), and [media (audio and video)](https://www.gemboxsoftware.com/presentation/examples/powerpoint-audio-video/406).
- Get, create, or edit the text in shapes and table cells specified through [paragraphs, runs, fields, and line breaks](https://www.gemboxsoftware.com/presentation/examples/powerpoint-textboxes/404).
- Get, create, or edit [hyperlinks](https://www.gemboxsoftware.com/presentation/examples/powerpoint-hyperlinks/409), [comments](https://www.gemboxsoftware.com/presentation/examples/powerpoint-comments/408), [placeholders](https://www.gemboxsoftware.com/presentation/examples/powerpoint-placeholders/402), [headers, and footers](https://www.gemboxsoftware.com/presentation/examples/powerpoint-header-footer/407).
- Get, create, or edit [shapes](https://www.gemboxsoftware.com/presentation/examples/powerpoint-shape-formatting/301), [table](https://www.gemboxsoftware.com/presentation/examples/powerpoint-table-formatting/602), [text box](https://www.gemboxsoftware.com/presentation/examples/powerpoint-textbox-formatting/302), [paragraph](https://www.gemboxsoftware.com/presentation/examples/powerpoint-paragraph-formatting/303), and [character](https://www.gemboxsoftware.com/presentation/examples/powerpoint-character-formatting/304) formatting.
- Get, create, or edit [table](https://www.gemboxsoftware.com/presentation/examples/powerpoint-table-styles/603) styles.
- Get and set [built-in and custom presentation properties](https://www.gemboxsoftware.com/presentation/examples/powerpoint-properties/410).
- Access or modify [slide size](https://www.gemboxsoftware.com/presentation/docs/GemBox.Presentation.PresentationDocument.html#GemBox_Presentation_PresentationDocument_SlideSize), [slide transition](https://www.gemboxsoftware.com/presentation/examples/powerpoint-slide-transition/501), and [slide show](https://www.gemboxsoftware.com/presentation/examples/powerpoint-slideshow/502) settings.
- [Preserve](https://www.gemboxsoftware.com/presentation/examples/powerpoint-diagrams/701) unsupported presentation content elements and properties when reading a presentation.

## Get Started

You are not sure how to start working with PowerPoint presentations in .NET using GemBox.Presentation? Check the code below that shows how to create a PPTX file from scratch and write 'Hello World!' on it.

```csharp
// If using Professional version, put your serial key below.
ComponentInfo.SetLicense("FREE-LIMITED-KEY");

// Create new empty presentation.
var presentation = new PresentationDocument();

// Add a new custom slide.
var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

// Add a rectangle and fill it with dark blue color.
var shape = slide.Content.AddShape(
    ShapeGeometryType.RoundedRectangle, 2, 2, 8, 4, LengthUnit.Centimeter);
shape.Format.Fill.SetSolid(Color.FromName(ColorName.DarkBlue));

// Add a paragraph and some text, and set text color to white.
var run = shape.Text.AddParagraph().AddRun("Hello World!");
run.Format.Fill.SetSolid(Color.FromName(ColorName.White));

// Save the presentation as PowerPoint's PPTX file.
presentation.Save("Writing.pptx");
```

For more GemBox.Presentation code examples and demos, please visit our [examples page](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-powerpoint-library/101).

## Installation

You can download GemBox.Presentation from [BugFixes 🛠️](https://www.gemboxsoftware.com/presentation/downloads/bugfixes.html) or from [NuGet 📦](https://www.nuget.org/packages/GemBox.Presentation/).

## Resources

- [Product Page](https://www.gemboxsoftware.com/presentation)
- [Examples](https://www.gemboxsoftware.com/presentation/examples)
- [Documentation](https://www.gemboxsoftware.com/presentation/docs/introduction.html)
- [API Reference](https://www.gemboxsoftware.com/presentation/docs/GemBox.Presentation.html)
- [Forum](https://forum.gemboxsoftware.com/c/gembox-presentation/8)
- [Blog](https://www.gemboxsoftware.com/gembox-presentation)
