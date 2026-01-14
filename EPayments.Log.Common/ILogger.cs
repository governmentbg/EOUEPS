using System;

namespace EPayments.Log
{
    public interface ILogger
    {
        void Log(LogLevel logLevel, string message);
        void Log(LogLevel logLevel, string message, Exception exception);
    }
}
