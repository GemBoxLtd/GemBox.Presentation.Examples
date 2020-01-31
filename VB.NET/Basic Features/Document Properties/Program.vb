Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("Reading.pptx")

        Dim slide = presentation.Slides(0)

        slide.Content.Drawings.Clear()

        ' Create "Built-in document properties" text box.
        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 0.5, 0.5, 12, 10, LengthUnit.Centimeter)
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkBlue))

        Dim paragraph = textBox.AddParagraph()
        paragraph.Format.Alignment = HorizontalAlignment.Left

        Dim run = paragraph.AddRun("Built-in document properties:")
        run.Format.Bold = True

        paragraph.AddLineBreak()

        For Each docProp In presentation.DocumentProperties.BuiltIn

            paragraph.AddRun(String.Format("{0}: {1}", docProp.Key, docProp.Value))
            paragraph.AddLineBreak()
        Next

        ' Create "Custom document properties" text box.
        textBox = slide.Content.AddTextBox(ShapeGeometryType.Rectangle, 14, 0.5, 12, 10, LengthUnit.Centimeter)
        textBox.Shape.Format.Outline.Fill.SetSolid(Color.FromName(ColorName.DarkBlue))

        paragraph = textBox.AddParagraph()
        paragraph.Format.Alignment = HorizontalAlignment.Left

        run = paragraph.AddRun("Custom document properties:")
        run.Format.Bold = True

        paragraph.AddLineBreak()

        For Each docProp In presentation.DocumentProperties.Custom

            paragraph.AddRun(String.Format("{0}: {1} (Type: {2})", docProp.Key, docProp.Value, docProp.Value.GetType()))
            paragraph.AddLineBreak()
        Next

        presentation.Save("Document Properties.pptx")
    End Sub
End Module