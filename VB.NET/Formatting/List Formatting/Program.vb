Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using the Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation As New PresentationDocument()
        Dim slide = presentation.Slides.AddNew(SlideLayoutType.Custom)

        ' Create number list items.
        Dim textBox = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 2, 2, 8, 5, LengthUnit.Centimeter)
        Dim paragraph = textBox.AddParagraph()
        paragraph.AddRun("First item.")
        paragraph.Format.List.NumberType = ListNumberType.DecimalPeriod
        paragraph.Format.List.Level = 0
        paragraph.Format.IndentationBeforeText = 27
        paragraph.Format.IndentationSpecial = -27

        paragraph = textBox.AddParagraph()
        paragraph.AddRun("Second item.")
        paragraph.Format.List.NumberType = ListNumberType.DecimalPeriod
        paragraph.Format.List.Level = 0
        paragraph.Format.IndentationBeforeText = 27
        paragraph.Format.IndentationSpecial = -27

        paragraph = textBox.AddParagraph()
        paragraph.AddRun("Second item's first sub-item.")
        paragraph.Format.List.NumberType = ListNumberType.LowerLetterPeriod
        paragraph.Format.List.Level = 1
        paragraph.Format.IndentationBeforeText = 54
        paragraph.Format.IndentationSpecial = -27

        paragraph = textBox.AddParagraph()
        paragraph.AddRun("Second item's second sub-item.")
        paragraph.Format.List.NumberType = ListNumberType.LowerLetterPeriod
        paragraph.Format.List.Level = 1
        paragraph.Format.IndentationBeforeText = 54
        paragraph.Format.IndentationSpecial = -27

        ' Create bullet list items.
        textBox = slide.Content.AddTextBox(ShapeGeometryType.RoundedRectangle, 2, 8, 8, 5, LengthUnit.Centimeter)
        paragraph = textBox.AddParagraph()
        paragraph.AddRun("First item.")
        paragraph.Format.List.BulletType = ListBulletType.FilledRound
        paragraph.Format.IndentationBeforeText = 27
        paragraph.Format.IndentationSpecial = -27

        paragraph = textBox.AddParagraph()
        paragraph.AddRun("Second item.")
        paragraph.Format.List.BulletType = ListBulletType.FilledRound
        paragraph.Format.IndentationBeforeText = 27
        paragraph.Format.IndentationSpecial = -27

        presentation.Save("Lists.pptx")
    End Sub
End Module
