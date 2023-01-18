Imports GemBox.Presentation
Imports GemBox.Presentation.Vba

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Example1()
        Example2()

    End Sub

    Sub Example1()
        Dim presentation As New PresentationDocument()
        presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create the module.
        Dim vbaModule As VbaModule = presentation.VbaProject.Modules.Add("SampleModule")
        vbaModule.Code =
"Sub SayHello()
    MsgBox ""Hello World!""
End Sub"

        ' Save the presentation as macro-enabled PowerPoint file.
        presentation.Save("AddVbaModule.pptm")
    End Sub

    Sub Example2()
        Dim presentation = PresentationDocument.Load("SampleVba.pptm")

        ' Get the module.
        Dim vbaModule As VbaModule = presentation.VbaProject.Modules("Slide1")
        ' Update text for the popup message.
        vbaModule.Code = vbaModule.Code.Replace("Hello world!", "Hello from GemBox.Presentation!")

        presentation.Save("UpdateVbaModule.pptm")
    End Sub

End Module