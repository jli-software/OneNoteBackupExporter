using System.ComponentModel;
using System.Runtime.InteropServices;

namespace OneNoteExporter.Helpers;

/// <summary>
/// Creates stable English error messages without exposing localized Windows or COM text.
/// </summary>
public static class UserFacingError
{
    public static string Describe(Exception exception, string message)
        => $"{message} ({GetTechnicalCode(exception)})";

    public static string GetTechnicalCode(Exception exception)
    {
        Exception relevantException = exception.InnerException ?? exception;

        return relevantException switch
        {
            COMException comException => $"COM 0x{comException.HResult:X8}",
            Win32Exception win32Exception => $"Windows error {win32Exception.NativeErrorCode}",
            _ => $"error 0x{relevantException.HResult:X8}"
        };
    }
}
