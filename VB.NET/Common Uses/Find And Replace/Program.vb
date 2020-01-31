Imports System.Linq
Imports GemBox.Presentation

Module Program

    Sub Main()

        ' If using Professional version, put your serial key below.
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")

        Dim presentation = PresentationDocument.Load("FindAndReplace.pptx")

        Dim slide = presentation.Slides(0)

        slide.TextContent.Replace("companyName", "Acme Corporation")

        Dim items As String()() = 
        {
            new String() {"January",  "$14.2M", "$1.6M", "$12.5M", "$2.3M"},
            new String() {"February", "$15.2M", "$0.6M", "$11.3M", "$4.3M"},
            new String() {"March",    "$17.2M", "$0M",   "$12.1M", "$52.3M"},
            new String() {"April",    "$7.2M",  "$1.6M", "$7.2M",  "$0M"},
            new String() {"May",      "$3.2M",  "$1.6M", "$0M",    "$3.2M"}
        }

        Dim table = slide.Content.Drawings.OfType(Of GraphicFrame).First().Table

        ' Clone the row with custom text tags (e.g. '@month', '@revenue', etc.).
        For  i = 0 to items.Length - 2
            table.Rows.AddClone(table.Rows(1))
        Next

        ' Replace custom text tags with data.
        for i = 0 to items.Length - 1
            Dim row = table.Rows(i + 1)
            Dim item = items(i)
            row.TextContent.Replace("@month", item(0))
            row.TextContent.Replace("@revenue", item(1))
            row.TextContent.Replace("@cashExpense", item(2))
            row.TextContent.Replace("@operatingIncome", item(3))
            row.TextContent.Replace("@operatingExpense", item(4))
        Next

        ' Find all text content with "0M" value in the table.
        Dim zeroValueRanges = table.TextContent.Find("0M")

        presentation.Save("Find And Replace.pptx")
    End Sub
End Module