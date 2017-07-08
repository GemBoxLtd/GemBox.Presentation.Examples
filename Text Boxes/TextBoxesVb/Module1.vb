Imports System
Imports GemBox.Presentation

Module Module1

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As PresentationDocument = New PresentationDocument

        ' Create New presentation slide.
        Dim slide As Slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create first text box.
        Dim textBox As TextBox = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 2, 2, 8, 8, LengthUnit.Centimeter)

        ' Set shape outline format.
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Create first paragraph with single run element.
        Dim run As TextRun = textBox.AddParagraph().AddRun("Shows how to use text boxes with GemBox.Presentation component.")
        run.Format.Bold = True

        ' Create empty paragraph.
        textBox.AddParagraph()

        ' Create (mixed-element) paragraph.
        Dim paragraph As TextParagraph = textBox.AddParagraph()

        ' Create And add a run element.
        run = paragraph.AddRun("Today's date: ")

        ' Create And add a "DateTime" text field element.
        Dim field As TextField = paragraph.AddField(TextFieldType.DateTime)

        ' Create empty paragraph.
        textBox.AddParagraph()

        ' Create (multi-line) paragraph.
        paragraph = textBox.AddParagraph()

        ' Create And add a first run element.
        run = paragraph.AddRun("This is a ...")

        ' Create And add a line break element.
        Dim lb As TextLineBreak = paragraph.AddLineBreak()

        ' Create And add a second run element.
        run = paragraph.AddRun("... multi-line paragraph.")

        ' Create second text box.
        textBox = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 12, 2, 8, 4, LengthUnit.Centimeter)

        ' Set shape outline format.
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkGray))

        ' Create a list.
        paragraph = textBox.AddParagraph()
        run = paragraph.AddRun("This is a paragraph list:")

        paragraph = textBox.AddParagraph()
        paragraph.Format.List.NumberType = ListNumberType.DecimalPeriod
        run = paragraph.AddRun("First list item")

        paragraph = textBox.AddParagraph()
        paragraph.Format.List.NumberType = ListNumberType.DecimalPeriod
        run = paragraph.AddRun("Second list item")

        paragraph = textBox.AddParagraph()
        paragraph.Format.List.NumberType = ListNumberType.DecimalPeriod
        run = paragraph.AddRun("Third list item")

        presentation.Save("Text Boxes.pptx")

    End Sub

End Module