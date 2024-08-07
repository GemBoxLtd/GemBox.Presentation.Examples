Imports GemBox.Presentation
Imports System

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")
        Dim properties = presentation.DocumentProperties

        Console.WriteLine("# Built-in document properties:")

        ' Write built-in document properties.
        properties.BuiltIn(BuiltInDocumentProperty.Title) = "My Title"
        properties.BuiltIn(BuiltInDocumentProperty.DateLastSaved) = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ")

        ' Read built-in document properties.
        For Each builtinProperty In properties.BuiltIn
            Console.WriteLine($"{builtinProperty.Key,20}: {builtinProperty.Value}")
        Next

        Console.WriteLine()
        Console.WriteLine("# Custom document properties:")

        ' Write custom document properties.
        properties.Custom("My Custom Property 1") = "My Custom Value"
        properties.Custom("My Custom Property 2") = 123.4

        ' Read custom document properties.
        For Each customProperty In properties.Custom
            Console.WriteLine($"{customProperty.Key,20}: {customProperty.Value,-20} [{customProperty.Value.GetType()}]")
        Next

    End Sub
End Module
