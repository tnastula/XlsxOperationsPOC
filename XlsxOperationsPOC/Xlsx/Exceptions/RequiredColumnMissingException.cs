namespace XlsxOperationsPOC.Xlsx.Exceptions;

public class RequiredColumnMissingException : Exception
{
    public RequiredColumnMissingException()
    {
    }

    public RequiredColumnMissingException(string? message) 
        : base(message)
    {
    }

    public RequiredColumnMissingException(string? message, Exception? innerException) 
        : base(message, innerException)
    {
        
    }
}