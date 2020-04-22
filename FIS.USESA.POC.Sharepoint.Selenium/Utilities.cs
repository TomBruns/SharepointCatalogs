using System;
using System.Collections.Generic;
using System.Text;

namespace FIS.USESA.POC.Sharepoint.Selenium
{
    internal static class Utilities
    {
        /// <summary>
        /// Write a log message to the console
        /// </summary>
        /// <param name="message"></param>
        internal static void WriteToConsole(string message)
        {
            ConsoleColor defaultColor = System.Console.ForegroundColor;
            System.Console.ForegroundColor = ConsoleColor.Blue;
            System.Console.WriteLine(message);
            System.Console.ForegroundColor = defaultColor;
        }
    }
}
