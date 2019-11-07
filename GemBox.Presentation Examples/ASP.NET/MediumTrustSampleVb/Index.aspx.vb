Imports System
Imports System.Data
Imports System.IO
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports GemBox.Presentation

Partial Public Class Index
    Inherits Page
    Protected Sub Page_Load(sender As Object, e As EventArgs)

        ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        AddHandler ComponentInfo.FreeLimitReached, Sub(s1, e1) e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial

        ' By specifying a location that Is under ASP.NET application's control, 
        ' GemBox.Presentation can retrieve font data when saving to PDF, even in Medium Trust environment.
        FontSettings.FontsBaseDirectory = Server.MapPath("Fonts/")

        If Not Me.Page.IsPostBack Then

            ' Fill grid view with some default data.
            Dim dataTable = New DataTable()
            dataTable.Columns.Add("Name", GetType(String))
            dataTable.Columns.Add("Estimated", GetType(String))
            dataTable.Columns.Add("Change", GetType(String))

            dataTable.Rows.Add("Revenues", "$14.2M", "(0.5%)")
            dataTable.Rows.Add("Cash Expense", "$1.6M", "0.7%")
            dataTable.Rows.Add("Operating Expense", "$12.5M", "0.3%")
            dataTable.Rows.Add("Operating Income", "$2.3M", "(0.2%)")

            Me.Session("highlightsDataTable") = dataTable

            Me.SetDataBinding()
        End If
    End Sub

    Private Shared Sub UpdatePresentation(presentation As PresentationDocument, title As String, summaryHeading As String, summaryBullets As String, highlightsHeading As String, highlightsDataTable As DataTable)

        ' Populate the first slide with data.
        Dim slide = presentation.Slides(0)
        Dim shape = DirectCast(slide.Content.Drawings(0), Shape)
        If Not String.IsNullOrEmpty(title) Then
            shape.Text.Paragraphs(0).AddRun(title.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
        End If

        ' Populate the second slide with data.
        slide = presentation.Slides(1)
        shape = DirectCast(slide.Content.Drawings(0), Shape)
        If Not String.IsNullOrEmpty(summaryHeading) Then
            shape.Text.Paragraphs(0).AddRun(summaryHeading.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
        End If
        shape = DirectCast(slide.Content.Drawings(1), Shape)
        shape.Text.Paragraphs.Clear()
        Dim summaryBulletLines = summaryBullets.Split(New String() {vbCr & vbLf}, StringSplitOptions.RemoveEmptyEntries)
        For Each summaryBulletLine In summaryBulletLines
            shape.Text.AddParagraph().AddRun(summaryBulletLine.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
        Next

        ' Populate the third slide with data.
        slide = presentation.Slides(2)
        shape = DirectCast(slide.Content.Drawings(0), Shape)
        If Not String.IsNullOrEmpty(highlightsHeading) Then
            shape.Text.Paragraphs(0).AddRun(highlightsHeading.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
        End If
        Dim frame = DirectCast(slide.Content.Drawings(1), GraphicFrame)
        Dim table = frame.Table
        For i As Integer = table.Rows.Count - 1 To 1 Step -1
            table.Rows.RemoveAt(i)
        Next
        For Each highlightDataRow As DataRow In highlightsDataTable.Rows
            Dim row = table.Rows.AddNew(table.Rows(0).Height)

            Dim cell = row.Cells.AddNew()
            Dim value = DirectCast(highlightDataRow("Name"), String)
            If Not String.IsNullOrEmpty(value) Then
                cell.Text.AddParagraph().AddRun(value.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
            End If

            cell = row.Cells.AddNew()
            value = DirectCast(highlightDataRow("Estimated"), String)
            If Not String.IsNullOrEmpty(value) Then
                cell.Text.AddParagraph().AddRun(value.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
            End If

            cell = row.Cells.AddNew()
            value = DirectCast(highlightDataRow("Change"), String)
            If Not String.IsNullOrEmpty(value) Then
                cell.Text.AddParagraph().AddRun(value.Replace(ControlChars.VerticalTab, " "c).Replace(ControlChars.Cr, " "c).Replace(ControlChars.Lf, " "c))
            End If
        Next
    End Sub

    Protected Sub generateButton_Click(sender As Object, e As EventArgs)

        Dim path_ As String = Path.Combine(Request.PhysicalApplicationPath, "Template.pptx")

        ' Load template presentation.
        Dim presentation = PresentationDocument.Load(path_)

        ' Populate the template presentation with data.
        UpdatePresentation(presentation, Me.titleTextBox.Text, Me.summaryHeadingTextBox.Text, Me.summaryBulletsTextBox.Text, Me.highlightsHeadingTextBox.Text, DirectCast(Session("highlightsDataTable"), DataTable))

        ' Stream the presentation to the browser.
        Dim fileName As String = "Presentation." + Me.outputFormatDropDownList.SelectedValue
        presentation.Save(Me.Response, fileName)
    End Sub

    Protected Sub addRowLinkButton_Click(sender As Object, e As EventArgs)

        Dim dataTable = DirectCast(Session("highlightsDataTable"), DataTable)
        dataTable.Rows.Add("", "", "")
        Me.SetDataBinding()
    End Sub

    Protected Sub highlightsGridView_RowEditing(sender As Object, e As GridViewEditEventArgs)

        Me.highlightsGridView.EditIndex = e.NewEditIndex
        Me.SetDataBinding()
    End Sub

    Protected Sub highlightsGridView_RowUpdating(sender As Object, e As GridViewUpdateEventArgs)

        Dim i As Integer
        Dim rowIndex As Integer = e.RowIndex
        Dim dataTable = DirectCast(Session("highlightsDataTable"), DataTable)

        For i = 0 To dataTable.Columns.Count - 1

            Dim editTextBox = TryCast(Me.highlightsGridView.Rows(rowIndex).Cells(i + 1).Controls(0), System.Web.UI.WebControls.TextBox)

            If editTextBox IsNot Nothing Then
                dataTable.Rows(rowIndex)(i) = editTextBox.Text
            End If
        Next

        Me.highlightsGridView.EditIndex = -1
        Me.SetDataBinding()
    End Sub

    Protected Sub highlightsGridView_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs)

        Me.highlightsGridView.EditIndex = -1
        Me.SetDataBinding()
    End Sub

    Protected Sub highlightsGridView_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)

        Dim dataTable = DirectCast(Session("highlightsDataTable"), DataTable)

        dataTable.Rows(e.RowIndex).Delete()
        Me.SetDataBinding()
    End Sub

    Private Sub SetDataBinding()

        Dim dataTable = DirectCast(Session("highlightsDataTable"), DataTable)

        Dim dataView As DataView = dataTable.DefaultView
        Me.highlightsGridView.DataSource = dataView
        dataView.AllowDelete = True

        Me.highlightsGridView.DataBind()
    End Sub

    Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function
End Class