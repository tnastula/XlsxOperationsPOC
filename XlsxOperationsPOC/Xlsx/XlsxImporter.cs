using System.Reflection;
using NPOI.SS.UserModel;
using XlsxOperationsPOC.Xlsx.Exceptions;
using XlsxOperationsPOC.Xlsx.Interfaces;

namespace XlsxOperationsPOC.Xlsx;

public class XlsxImporter<T>
    where T : IXlsxDataObject
{
    public List<T> Data { get; private set; }
    private List<ColumnHeader> Headers { get; set; }

    public XlsxImporter()
    {
        T? dataObject = (T?)Activator.CreateInstance(typeof(T), false);
        if (dataObject == null)
        {
            throw new DataObjectCreationException($"Failed to create object of {typeof(T)} type.");
        }

        Headers = dataObject.GetColumnHeaders();
        Data = new();
    }

    public void Import(string fileName)
    {
        var workbook = WorkbookFactory.Create(fileName);
        var sheet = workbook.GetSheetAt(0);
        var headerRow = sheet.GetRow(0);
        ValidateHeaderRow(headerRow);

        Data = new(sheet.LastRowNum - 1);
        for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                continue;
            }

            T? dataObject = (T?)Activator.CreateInstance(typeof(T), false);
            if (dataObject == null)
            {
                throw new DataObjectCreationException($"Failed to create object of {typeof(T)} type.");
            }

            ReadRow(row, dataObject);
        }
    }

    private void ReadRow(IRow row, T dataObject)
    {
        foreach (ColumnHeader header in Headers)
        {
            if (header.ColumnIndex == null)
            {
                continue;
            }

            ICell cell = row.GetCell(header.ColumnIndex.GetValueOrDefault());
            if (cell == null || cell.CellType == CellType.Blank)
            {
                continue;
            }

            ReadCellValue(cell, header.CodePropertyName, dataObject);
        }

        Data.Add(dataObject);
    }

    private void ReadCellValue(ICell cell, string propertyName, T dataObject)
    {
        PropertyInfo propertyInfo = dataObject.GetType().GetProperty(propertyName)
                                    ?? throw new InvalidOperationException(
                                        $"Failed to obtain PropertyInfo for {propertyName} field");

        if (propertyInfo.PropertyType == typeof(int))
        {
            propertyInfo.SetValue(dataObject, (int)cell.NumericCellValue);
        }
        else if (propertyInfo.PropertyType == typeof(double))
        {
            propertyInfo.SetValue(dataObject, cell.NumericCellValue);
        }
        else if (propertyInfo.PropertyType == typeof(decimal))
        {
            propertyInfo.SetValue(dataObject, (decimal)cell.NumericCellValue);
        }
        else if (propertyInfo.PropertyType == typeof(DateTime))
        {
            propertyInfo.SetValue(dataObject, cell.DateCellValue);
        }
        else if (propertyInfo.PropertyType == typeof(string))
        {
            propertyInfo.SetValue(dataObject, cell.StringCellValue);
        }
        else if (propertyInfo.PropertyType == typeof(bool))
        {
            propertyInfo.SetValue(dataObject, cell.BooleanCellValue);
        }
    }

    private void ValidateHeaderRow(IRow row)
    {
        Headers.ForEach(x => x.ColumnIndex = null);

        foreach (var headerCell in row.Cells)
        {
            if (headerCell.CellType != CellType.String)
            {
                continue;
            }

            ColumnHeader? header = Headers.Find(x => x.SpreadsheetColumnName == headerCell.StringCellValue);
            if (header == null)
            {
                continue;
            }

            if (header.ColumnIndex != null)
            {
                throw new DuplicateColumnException($"Column {header.SpreadsheetColumnName} occurs at least twice.");
            }

            header.ColumnIndex = headerCell.ColumnIndex;
        }

        List<ColumnHeader> missingRequiredHeaders = Headers
            .Where(x => x.ColumnIndex == null && x.IsRequired)
            .ToList();

        if (missingRequiredHeaders.Count > 0)
        {
            string missingHeadersString = String.Join(
                ", ",
                missingRequiredHeaders.Select(x => $"'{x.SpreadsheetColumnName}'"));

            throw new RequiredColumnMissingException(
                $"The following required columns are missing: {missingHeadersString}");
        }
    }
}