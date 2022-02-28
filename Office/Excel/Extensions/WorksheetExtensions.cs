using DocumentFormat.OpenXml.Spreadsheet;

namespace Elfland.Office.Excel.Extensions;

public static class WorksheetExtensions
{
    public static Cell GetOrCreateCell(this Worksheet worksheet, string columnName, int rowIndex)
    {
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = Convert.ToUInt32(rowIndex) };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.
        if (
            row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count()
            > 0
        )
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);
            return newCell;
        }
    }
}
