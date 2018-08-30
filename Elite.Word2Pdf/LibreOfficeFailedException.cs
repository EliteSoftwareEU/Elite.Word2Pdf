using System;
namespace Elite.Word2Pdf
{
    public class LibreOfficeFailedException : Exception
    {
        public LibreOfficeFailedException(int exitCode)
            : base(string.Format("LibreOffice has failed with code: {}", exitCode))
        { }
    }

}
