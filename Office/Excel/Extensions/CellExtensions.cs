using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Elfland.Office.Excel.Extensions;

public static class CellExtensions
{
    public static void UpdateValue(
        this Cell cell,
        string cellValue,
        CellValues cellType = CellValues.String
    )
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(cellType);
    }

    public static void UpdateValue(this Cell cell, bool cellValue)
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
    }

    public static void UpdateValue(this Cell cell, int cellValue)
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
    }

    public static void UpdateValue(this Cell cell, double cellValue)
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
    }

    public static void UpdateValue(this Cell cell, decimal cellValue)
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.Number);
    }

    public static void UpdateValue(this Cell cell, DateTime cellValue)
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.Date);
    }

    public static void UpdateValue(this Cell cell, DateTimeOffset cellValue)
    {
        cell.CellValue = new CellValue(cellValue);
        cell.DataType = new EnumValue<CellValues>(CellValues.Date);
    }
}
