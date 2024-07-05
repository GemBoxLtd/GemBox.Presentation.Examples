[![NuGet version](https://img.shields.io/nuget/v/GemBox.Presentation?style=for-the-badge)](https://www.nuget.org/packages/GemBox.Presentation/) [![NuGet downloads](https://img.shields.io/nuget/dt/GemBox.Presentation?style=for-the-badge)](https://www.nuget.org/packages/GemBox.Presentation/) [![Visual Studio Marketplace rating](https://img.shields.io/visual-studio-marketplace/stars/GemBoxSoftware.GemBoxPresentation?style=for-the-badge)](https://marketplace.visualstudio.com/items?itemName=GemBoxSoftware.GemBoxPresentation)

## What is GemBox.Presentation?

GemBox.Presentation is a .NET component that enables you to read, write, convert, and print presentation files (PPTX, PPT, and PPSX) from .NET applications.

With GemBox.Presentation you get a fast and reliable component that’s easy to use and doesn't depend on Microsoft PowerPoint. It requires only .NET and it’s much faster than Microsoft Office Automation!


## GemBox.Presentation Features

- [Read](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-open-read-powerpoint/201) PPT and PPTX presentations.
- [Create or write](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-create-write-powerpoint/202) and [convert](https://www.gemboxsoftware.com/presentation/examples/c-sharp-convert-powerpoint-to-pdf/204) to PPT, PPTX, PDF, XPS, and image formats (SVG, PNG, JPEG, GIF, BMP, TIFF, WMP).
- View presentations in [Azure Functions](https://www.gemboxsoftware.com/presentation/examples/create-powerpoint-pdf-on-azure-functions-app-service/2201), [Blazor](https://www.gemboxsoftware.com/presentation/examples/blazor-create-powerpoint/2002), [ASP.NET Core](https://www.gemboxsoftware.com/presentation/examples/asp-net-core-create-powerpoint-pptx-pdf/2001), [ASP.NET](https://www.gemboxsoftware.com/presentation/examples/asp-net-powerpoint-export/1601), [MAUI](https://www.gemboxsoftware.com/presentation/examples/create-powerpoint-file-maui/2102), and [WPF]((https://www.gemboxsoftware.com/presentation/examples/powerpoint-xpsdocument-wpf/1701) applications.
- Process presentations on Windows, [Linux, macOS](https://www.gemboxsoftware.com/presentation/examples/create-powerpoint-pdf-on-linux-net-core/1901), [Android, and iOS](https://www.gemboxsoftware.com/presentation/examples/create-powerpoint-file-xamarin/2101) operating systems.
- [Print](https://www.gemboxsoftware.com/presentation/examples/c-sharp-print-powerpoint/251) presentations.
- [Protect](https://www.gemboxsoftware.com/presentation/examples/pptx-modify-protection/804), [encrypt](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-pptx-encryption/803), and [digitally sign](https://www.gemboxsoftware.com/presentation/examples/pptx-digital-signature/805) presentations.
- Get, create, or edit [master slides, layout slides, slides](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-powerpoint-slides/401), [master notes slide, and notes slides](https://www.gemboxsoftware.com/presentation/examples/c-sharp-vb-net-powerpoint-slide-notes/411).
- Get, create, or edit drawings like [text boxes](https://www.gemboxsoftware.com/presentation/examples/powerpoint-textboxes/404), [connectors](https://www.gemboxsoftware.com/presentation/examples/powerpoint-shapes/403), [pictures](https://www.gemboxsoftware.com/presentation/examples/powerpoint-pictures/405), [tables](https://www.gemboxsoftware.com/presentation/examples/powerpoint-tables/601), [charts](https://www.gemboxsoftware.com/presentation/examples/powerpoint-charts/412), and [media (audio and video)](https://www.gemboxsoftware.com/presentation/examples/powerpoint-audio-video/406).
- Get, create, or edit text in shapes and table cells specified through [paragraphs, runs, fields, and line breaks](https://www.gemboxsoftware.com/presentation/examples/powerpoint-textboxes/404).
- Get, create, or edit [hyperlinks](https://www.gemboxsoftware.com/presentation/examples/powerpoint-hyperlinks/409), [comments](https://www.gemboxsoftware.com/presentation/examples/powerpoint-comments/408), [placeholders](https://www.gemboxsoftware.com/presentation/examples/powerpoint-placeholders/402), [headers, and footers](https://www.gemboxsoftware.com/presentation/examples/powerpoint-header-footer/407).
- Get, create, or edit [shape](https://www.gemboxsoftware.com/presentation/examples/powerpoint-shape-formatting/301), [table](https://www.gemboxsoftware.com/presentation/examples/powerpoint-table-formatting/602), [text box](https://www.gemboxsoftware.com/presentation/examples/powerpoint-textbox-formatting/302), [paragraph](https://www.gemboxsoftware.com/presentation/examples/powerpoint-paragraph-formatting/303), and [character](https://www.gemboxsoftware.com/presentation/examples/powerpoint-character-formatting/304) formatting.
- Get, create or edit [table](https://www.gemboxsoftware.com/presentation/examples/powerpoint-table-styles/603) styles.
- Get and set [built-in and custom presentation properties](https://www.gemboxsoftware.com/presentation/examples/powerpoint-properties/410).
- Load [HTML](https://www.gemboxsoftware.com/presentation/examples/powerpoint-load-html/208) content into presentations.
- Access or modify [slide size](https://www.gemboxsoftware.com/presentation/docs/GemBox.Presentation.PresentationDocument.html#GemBox_Presentation_PresentationDocument_SlideSize), [slide transition](https://www.gemboxsoftware.com/presentation/examples/powerpoint-slide-transition/501), [slide show](https://www.gemboxsoftware.com/presentation/examples/powerpoint-slideshow/502), [macros](https://www.gemboxsoftware.com/presentation/examples/vba-macros/506), and more.
- [Specify fonts location](https://www.gemboxsoftware.com/presentation/examples/private-fonts/503) when exporting to PDF, XPS, or image formats.
- [Preserve](https://www.gemboxsoftware.com/presentation/examples/powerpoint-preservation/701) unsupported presentation content elements and properties when reading a presentation.
- [Medium trust](https://www.gemboxsoftware.com/presentation/examples/asp-net-powerpoint-export/1601) support.

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

You can download GemBox.Presentation from [NuGet 📦](https://www.nuget.org/packages/GemBox.Presentation/) or from [BugFixes 🛠️](https://www.gemboxsoftware.com/presentation/downloads/bugfixes.html).

## Resources

- [Product Page](https://www.gemboxsoftware.com/presentation)
- [Examples](https://www.gemboxsoftware.com/presentation/examples)
- [Documentation](https://www.gemboxsoftware.com/presentation/docs/introduction.html)
- [API Reference](https://www.gemboxsoftware.com/presentation/docs/GemBox.Presentation.html)
- [Forum](https://forum.gemboxsoftware.com/c/gembox-presentation/8)
- [Blog](https://www.gemboxsoftware.com/gembox-presentation)
