namespace XlsxOperationsPOC.Xlsx.Exceptions;

public class UnspecifiedColumnIndexException : Exception
{
    public UnspecifiedColumnIndexException()
    {
    }

    public UnspecifiedColumnIndexException(string? message)
        : base(message)
    {
    }

    public UnspecifiedColumnIndexException(string? message, Exception? innerException)
        : base(message, innerException)
    {
    }
}