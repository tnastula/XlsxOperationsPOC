namespace XlsxOperationsPOC.Xlsx;

public class ColumnHeader
{
    public int? ColumnIndex { get; set; }
    
    public readonly string SpreadsheetColumnName;
    public readonly string CodePropertyName;
    public readonly bool IsRequired;

    public ColumnHeader(string spreadsheetColumnName, string codePropertyName, bool isRequired, int columnIndex)
    {
        SpreadsheetColumnName = spreadsheetColumnName;
        CodePropertyName = codePropertyName;
        IsRequired = isRequired;
        ColumnIndex = columnIndex;
    }
}