Imports System
Imports System.Linq
Imports System.Text
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim sb = New StringBuilder()

        Dim slide = presentation.Slides(0)

        For Each shape In slide.Content.Drawings.OfType(Of Shape)

            sb.AppendFormat("Shape ShapeType={0}:", shape.ShapeType)
            sb.AppendLine()

            For Each paragraph In shape.Text.Paragraphs

                For Each run In paragraph.Elements.OfType(Of TextRun)

                    Dim isBold = run.Format.Bold
                    Dim text = run.Text

                    sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
                Next

                sb.AppendLine()
            Next

            sb.AppendLine("----------")
        Next

        Console.WriteLine(sb.ToString())
    End Sub
End Module