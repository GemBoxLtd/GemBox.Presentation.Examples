using System;
using System.Data;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using GemBox.Presentation;

public partial class Index : Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        ComponentInfo.FreeLimitReached += (s1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

        if (!this.Page.IsPostBack)
        {
            // Fill grid view with some default data.
            var dataTable = new DataTable();
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Estimated", typeof(string));
            dataTable.Columns.Add("Change", typeof(string));

            dataTable.Rows.Add("Revenues", "$14.2M", "(0.5%)");
            dataTable.Rows.Add("Cash Expense", "$1.6M", "0.7%");
            dataTable.Rows.Add("Operating Expense", "$12.5M", "0.3%");
            dataTable.Rows.Add("Operating Income", "$2.3M", "(0.2%)");

            this.Session["highlightsDataTable"] = dataTable;

            this.SetDataBinding();
        }
    }

    private static void UpdatePresentation(PresentationDocument presentation, string title, string summaryHeading, string summaryBullets, string highlightsHeading, DataTable highlightsDataTable)
    {
        // Populate the first slide with data.
        var slide = presentation.Slides[0];
        var shape = (Shape)slide.Content.Drawings[0];
        if (!string.IsNullOrEmpty(title))
            shape.Text.Paragraphs[0].AddRun(title.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));

        // Populate the second slide with data.
        slide = presentation.Slides[1];
        shape = (Shape)slide.Content.Drawings[0];
        if (!string.IsNullOrEmpty(summaryHeading))
            shape.Text.Paragraphs[0].AddRun(summaryHeading.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));
        shape = (Shape)slide.Content.Drawings[1];
        shape.Text.Paragraphs.Clear();
        var summaryBulletLines = summaryBullets.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var summaryBulletLine in summaryBulletLines)
            shape.Text.AddParagraph().AddRun(summaryBulletLine.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));

        // Populate the third slide with data.
        slide = presentation.Slides[2];
        shape = (Shape)slide.Content.Drawings[0];
        if (!string.IsNullOrEmpty(highlightsHeading))
            shape.Text.Paragraphs[0].AddRun(highlightsHeading.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));
        var frame = (GraphicFrame)slide.Content.Drawings[1];
        var table = frame.Table;
        for (int i = table.Rows.Count - 1; i > 0; --i)
            table.Rows.RemoveAt(i);
        foreach (DataRow highlightDataRow in highlightsDataTable.Rows)
        {
            var row = table.Rows.AddNew(table.Rows[0].Height);

            var cell = row.Cells.AddNew();
            var value = (string)highlightDataRow["Name"];
            if (!string.IsNullOrEmpty(value))
                cell.Text.AddParagraph().AddRun(value.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));

            cell = row.Cells.AddNew();
            value = (string)highlightDataRow["Estimated"];
            if (!string.IsNullOrEmpty(value))
                cell.Text.AddParagraph().AddRun(value.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));

            cell = row.Cells.AddNew();
            value = (string)highlightDataRow["Change"];
            if (!string.IsNullOrEmpty(value))
                cell.Text.AddParagraph().AddRun(value.Replace('\v', ' ').Replace('\r', ' ').Replace('\n', ' '));
        }
    }

    protected void generateButton_Click(object sender, EventArgs e)
    {
        string path = Path.Combine(Request.PhysicalApplicationPath, "Template.pptx");

        // Load template presentation.
        var presentation = PresentationDocument.Load(path);

        // Populate the template presentation with data.
        UpdatePresentation(presentation, this.titleTextBox.Text, this.summaryHeadingTextBox.Text, this.summaryBulletsTextBox.Text, this.highlightsHeadingTextBox.Text, (DataTable)Session["highlightsDataTable"]);

        // Stream the presentation to the browser.
        string fileName = "Presentation." + this.outputFormatDropDownList.SelectedValue;
        presentation.Save(this.Response, fileName);
    }

    protected void addRowLinkButton_Click(object sender, EventArgs e)
    {
        var dataTable = (DataTable)Session["highlightsDataTable"];
        dataTable.Rows.Add("", "", "");
        this.SetDataBinding();
    }

    protected void highlightsGridView_RowEditing(object sender, GridViewEditEventArgs e)
    {
        this.highlightsGridView.EditIndex = e.NewEditIndex;
        this.SetDataBinding();
    }

    protected void highlightsGridView_RowUpdating(object sender, GridViewUpdateEventArgs e)
    {
        int i;
        int rowIndex = e.RowIndex;
        var dataTable = (DataTable)Session["highlightsDataTable"];

        for (i = 0; i < dataTable.Columns.Count; i++)
        {
            var editTextBox = this.highlightsGridView.Rows[rowIndex].Cells[i + 1].Controls[0] as System.Web.UI.WebControls.TextBox;

            if (editTextBox != null)
                dataTable.Rows[rowIndex][i] = editTextBox.Text;
        }

        this.highlightsGridView.EditIndex = -1;
        this.SetDataBinding();
    }

    protected void highlightsGridView_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
    {
        this.highlightsGridView.EditIndex = -1;
        this.SetDataBinding();
    }

    protected void highlightsGridView_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        var dataTable = (DataTable)Session["highlightsDataTable"];

        dataTable.Rows[e.RowIndex].Delete();
        this.SetDataBinding();
    }

    private void SetDataBinding()
    {
        var dataTable = (DataTable)Session["highlightsDataTable"];

        DataView dataView = dataTable.DefaultView;
        this.highlightsGridView.DataSource = dataView;
        dataView.AllowDelete = true;

        this.highlightsGridView.DataBind();
    }
}