using XlsxOperationsPOC.Xlsx;
using XlsxOperationsPOC.Xlsx.Interfaces;

namespace XlsxOperationsPOC;

public class FooBar : IXlsxDataObject
{
    public int CoolNumber { get; set; }
    public string? AwesomeText { get; set; }
    public DateTime AmazingDate { get; set; }

    public override string ToString()
    {
        return $"{CoolNumber}\t{AwesomeText}\t{AmazingDate.ToShortDateString()}";
    }

    public List<ColumnHeader> GetColumnHeaders()
    {
        return new List<ColumnHeader>
        {
            new("Cool number", nameof(CoolNumber), true, 0),
            new("Awesome text", nameof(AwesomeText), false, 1),
            new("Amazing date", nameof(AmazingDate), true, 2)
        };
    }
}