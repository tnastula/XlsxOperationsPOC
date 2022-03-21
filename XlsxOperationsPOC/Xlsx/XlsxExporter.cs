using System.Reflection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using XlsxOperationsPOC.Xlsx.Exceptions;
using XlsxOperationsPOC.Xlsx.Interfaces;

namespace XlsxOperationsPOC.Xlsx;

public class XlsxExporter<T>
    where T : IXlsxDataObject
{
    private List<T> Data { get; set; }
    private List<ColumnHeader> Headers { get; set; }
    private ICellStyle? DateStyle { get; set; }
    private ICellStyle? NumberStyle { get; set; }

    public XlsxExporter(List<T> data)
    {
        T? dataObject = (T?)Activator.CreateInstance(typeof(T), false);
        if (dataObject == null)
        {
            throw new DataObjectCreationException($"Failed to create object of {typeof(T)} type.");
        }

        Headers = dataObject.GetColumnHeaders();
        Data = data;
    }

    public void Export(string fileName)
    {
        var workbook = new XSSFWorkbook();
        InitializeCellStyles(workbook);

        var worksheet = workbook.CreateSheet("Worksheet");
        ValidateHeaders();
        WriteHeaderRow(worksheet);

        int currentRowNum = 1;
        while (currentRowNum <= Data.Count)
        {
            var currentRow = worksheet.CreateRow(currentRowNum);
            WriteRow(Data[currentRowNum - 1], currentRow);

            currentRowNum++;
        }

        using FileStream fileStream = new FileStream(fileName, FileMode.Create);
        workbook.Write(fileStream);
        fileStream.Close();
    }

    private void InitializeCellStyles(XSSFWorkbook workbook)
    {
        IDataFormat dataFormat = workbook.CreateDataFormat();

        DateStyle = workbook.CreateCellStyle();
        DateStyle.DataFormat = dataFormat.GetFormat("m/d/yy");

        NumberStyle = workbook.CreateCellStyle();
        NumberStyle.DataFormat = dataFormat.GetFormat("0");
    }

    private void ValidateHeaders()
    {
        if (Headers.Any(x => x.ColumnIndex == null))
        {
            throw new UnspecifiedColumnIndexException($"Detected at least one column without null index.");
        }

        if (Headers.Select(x => x.ColumnIndex).Distinct().Count() != Headers.Count)
        {
            throw new DuplicateColumnException($"Detected two or more columns occupying the same index.");
        }

        Headers = Headers.OrderBy(x => x.ColumnIndex).ToList();
    }

    private void WriteHeaderRow(ISheet worksheet)
    {
        var headerRow = worksheet.CreateRow(0);

        foreach (ColumnHeader header in Headers)
        {
            ICell headerCell = headerRow.CreateCell(header.ColumnIndex.GetValueOrDefault());
            headerCell.SetCellValue(header.SpreadsheetColumnName);
        }
    }

    private void WriteRow(T dataObject, IRow row)
    {
        foreach (ColumnHeader header in Headers)
        {
            PropertyInfo propertyInfo = dataObject.GetType().GetProperty(header.CodePropertyName)
                                        ?? throw new InvalidOperationException(
                                            $"Failed to obtain PropertyInfo for {header.CodePropertyName} field");
            
            ICell cell = row.CreateCell(header.ColumnIndex.GetValueOrDefault());
            object? cellValue = propertyInfo.GetValue(dataObject);
            if (cellValue == null)
            {
                continue;
            }
            
            if (propertyInfo.PropertyType == typeof(int))
            {
                cell.SetCellValue((int)cellValue);
                cell.CellStyle = NumberStyle;
            }
            else if (propertyInfo.PropertyType == typeof(double))
            {
                cell.SetCellValue((double)cellValue);
                cell.CellStyle = NumberStyle;
            }
            else if (propertyInfo.PropertyType == typeof(decimal))
            {
                cell.SetCellValue((double)cellValue);
                cell.CellStyle = NumberStyle;
            }
            else if (propertyInfo.PropertyType == typeof(DateTime))
            {
                cell.SetCellValue((DateTime)cellValue);
                cell.CellStyle = DateStyle;
            }
            else if (propertyInfo.PropertyType == typeof(string))
            {
                cell.SetCellValue((string)cellValue);
            }
            else if (propertyInfo.PropertyType == typeof(bool))
            {
                cell.SetCellValue((bool)cellValue);
            }
        }
    }
}