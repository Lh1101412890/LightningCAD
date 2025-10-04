using System;
using System.IO;

namespace LightningCAD.LightningExtension
{
    public static class ExceptionExtension
    {
        public static void Record(this Exception exception)
        {
            string errorMessage;
            if (exception.InnerException == null)
            {
                errorMessage = $"{DateTime.Now}:\n--{exception.Message}\n{exception.StackTrace}\n\n";
            }
            else
            {
                errorMessage = $"{DateTime.Now}:\n--{exception.Message}\n{exception.StackTrace}\n--{exception.InnerException.Message}\n{exception.InnerException.StackTrace}\n\n";
            }
            File.AppendAllText(Information.God.ErrorLog, errorMessage);
        }
    }
}