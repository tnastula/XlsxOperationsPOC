using XlsxOperationsPOC;
using XlsxOperationsPOC.Xlsx;

//
// WRITE TEST
//
List<FooBar> dataToWrite = new List<FooBar>
{
    new()
    {
        AmazingDate = new DateTime(2020, 2, 13),
        AwesomeText = "An awesome day",
        CoolNumber = 7
    },
    new()
    {
        AmazingDate = new DateTime(1987, 2, 26),
        AwesomeText = "What a cool day",
        CoolNumber = 33
    }
};

DateTime now = DateTime.Now;
string fileName = $"writtenOn-{now.ToShortDateString()}-{now.ToShortTimeString()}.xlsx";
fileName = fileName.Replace(':', ';');
var writeTest = new XlsxExporter<FooBar>(dataToWrite);
writeTest.Export(fileName);


//
// READ TEST
//
var readTest = new XlsxImporter<FooBar>();
readTest.Import("read-test.xlsx");
foreach (var data in readTest.Data)
{
    Console.WriteLine(data.ToString());
}