Imports System
Imports System.Linq
Imports System.Text
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = PresentationDocument.Load("Reading.pptx")

        Dim sb As New StringBuilder()

        Dim slide As Slide = presentation.Slides(0)

        For Each shape As Shape In slide.Content.Drawings.OfType(Of Shape)
            sb.AppendFormat("Shape ShapeType={0}:", shape.ShapeType)
            sb.AppendLine()

            For Each paragraph As TextParagraph In shape.Text.Paragraphs
                For Each run As TextRun In paragraph.Elements.OfType(Of TextRun)
                    Dim isBold As Boolean = run.Format.Bold
                    Dim text As String = run.Text

                    sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
                Next

                sb.AppendLine()
            Next

            sb.AppendLine("----------")
        Next

        Console.WriteLine(sb.ToString())

    End Sub

End Module