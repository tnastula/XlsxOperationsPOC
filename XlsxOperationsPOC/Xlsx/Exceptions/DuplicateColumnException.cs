namespace XlsxOperationsPOC.Xlsx.Exceptions;

public class DuplicateColumnException : Exception
{
    public DuplicateColumnException()
    {
    }

    public DuplicateColumnException(string? message) 
        : base(message)
    {
    }

    public DuplicateColumnException(string? message, Exception? innerException) 
        : base(message, innerException)
    {
        
    }
}