Imports System
Imports System.Text
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = PresentationDocument.Load("Reading.pptx")

        Dim sb As New StringBuilder()

        sb.AppendLine("Slide size (width X height):")

        Dim width As Length = presentation.SlideSize.Width
        Dim height As Length = presentation.SlideSize.Height

        For Each unit As LengthUnit In [Enum].GetValues(GetType(LengthUnit))

            sb.AppendFormat("{0} X {1} {2}",
                            width.To(unit),
                            height.To(unit),
                            unit.ToString().ToLowerInvariant())

            sb.AppendLine()

        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module