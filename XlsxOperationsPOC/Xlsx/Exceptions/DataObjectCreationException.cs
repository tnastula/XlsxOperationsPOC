namespace XlsxOperationsPOC.Xlsx.Exceptions;

public class DataObjectCreationException : Exception
{
    public DataObjectCreationException()
    {
    }

    public DataObjectCreationException(string? message) 
        : base(message)
    {
    }

    public DataObjectCreationException(string? message, Exception? innerException) 
        : base(message, innerException)
    {
    }
}