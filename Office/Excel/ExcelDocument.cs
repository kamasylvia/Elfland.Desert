using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Elfland.Office.Excel;

public class ExcelDocument : IDisposable
{
    private SpreadsheetDocument _spreadsheetDocument;

    /// <summary>
    /// Open or create a new Excel document. It supports creating from a template.
    /// </summary>
    /// <param name="filepath"></param>
    public ExcelDocument(string filepath)
    {
        if (string.IsNullOrEmpty(filepath))
        {
            throw new ArgumentException(
                $"'{nameof(filepath)}' cannot be null or empty.",
                nameof(filepath)
            );
        }

        _spreadsheetDocument = SpreadsheetDocument.CreateFromTemplate(filepath);
        InitializeSpreadsheetDocument();
    }

    private void InitializeSpreadsheetDocument()
    {
        // Add a WorkbookPart to the document.
        WorkbookPart workbookpart = _spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        // Add Sheets to the Workbook.
        Sheets sheets = _spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(
            new Sheets()
        );
    }

    public Worksheet InsertWorksheet(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            throw new ArgumentException(
                $"'{nameof(sheetName)}' cannot be null or empty.",
                nameof(sheetName)
            );
        }

        // Add a WorksheetPart to the WorkbookPart.
        var sheets = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
        WorksheetPart worksheetPart = _spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheet = new Sheet()
        {
            Id = _spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = Convert.ToUInt32(sheets.Count()) + 1,
            Name = sheetName
        };

        sheets.Append(sheet);

        _spreadsheetDocument.WorkbookPart.Workbook.Save();

        return worksheetPart.Worksheet;
    }

    public Worksheet GetWorksheet(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            throw new ArgumentException(
                $"'{nameof(sheetName)}' cannot be null or empty.",
                nameof(sheetName)
            );
        }

        var sheet = _spreadsheetDocument.WorkbookPart.Workbook
            .GetFirstChild<Sheets>()
            .Elements<Sheet>()
            .Where(s => s.Name == sheetName)
            .FirstOrDefault();

        var worksheetPart =
            _spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;

        return worksheetPart.Worksheet;
    }

    private bool _disposed = false;

    ~ExcelDocument() => Dispose(false);

    public void Dispose()
    {
        // Dispose of unmanaged resources.
        Dispose(true);
        // Suppress finalization.
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed || _spreadsheetDocument is null)
        {
            return;
        }

        // Dispose of managed resources here.
        if (disposing) { }

        // Dispose of any unmanaged resources not wrapped in safe handles.
        _spreadsheetDocument?.Close();
        _spreadsheetDocument = null;

        _disposed = true;
    }
}
