Imports System
Imports System.Text
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim sb = New StringBuilder()

        sb.AppendLine("Slide size (width X height):")

        Dim width = presentation.SlideSize.Width
        Dim height = presentation.SlideSize.Height

        For Each unit As LengthUnit In [Enum].GetValues(GetType(LengthUnit))

            sb.AppendFormat(
                "{0} X {1} {2}",
                width.To(unit),
                height.To(unit),
                unit.ToString().ToLowerInvariant())

            sb.AppendLine()
        Next

        Console.WriteLine(sb.ToString())
    End Sub
End Module