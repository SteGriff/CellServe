using System;
using System.Runtime.Serialization;

namespace CellServe.ExcelHandler
{
    public class CellServeException : ApplicationException
    {
        public CellServeException()
        {
        }

        public CellServeException(string message) : base(message)
        {
        }

        public CellServeException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected CellServeException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
