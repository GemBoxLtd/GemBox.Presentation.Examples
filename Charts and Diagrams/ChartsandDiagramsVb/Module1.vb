Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = PresentationDocument.Load("ChartsAndDiagrams.pptx")

        presentation.Save("Charts and Diagrams.pptx")

    End Sub

End Module