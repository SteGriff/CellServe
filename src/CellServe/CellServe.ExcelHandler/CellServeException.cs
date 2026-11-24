namespace CellServe.ExcelHandler;

public sealed class CellServeException : Exception
{
    public CellServeException(string message) : base(message)
    {
    }

    public CellServeException(string message, Exception innerException) : base(message, innerException)
    {
    }
}
