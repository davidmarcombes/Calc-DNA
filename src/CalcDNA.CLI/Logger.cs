using System;
using System.Collections.Generic;
using System.Text;

namespace CalcDNA.CLI
{
    /// <summary>
    /// Simple logger for console output
    /// </summary>
    public class Logger
    {
        // Note: simplistic logger for console output, function no static on purpose to allow future enhancements

        private void WriteLine(string message, ConsoleColor color)
        {
            try
            {
                Console.ForegroundColor = color;
                Console.WriteLine(message);
            }
            finally
            {
                Console.ResetColor();
            }
        }

        /// <summary>
        /// Log an info message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void Info(string message)
        {
            WriteLine(message, ConsoleColor.White);
        }

        /// <summary>
        /// Log an error message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void Error(string message)
        {
            WriteLine(message, ConsoleColor.Red);
        }

        /// <summary>
        /// Log a warning message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void Warning(string message)
        {
            WriteLine(message, ConsoleColor.Yellow);
        }

        /// <summary>
        /// Log a debug message
        /// </summary>
        /// <param name="message">The message to log</param>
        /// <param name="verbose">Whether to log the message</param>
        public void Debug(string message, bool verbose)
        {
            if (verbose)
            {
                WriteLine(message, ConsoleColor.DarkGray);
            }
        }

        /// <summary>
        /// Log a success message
        /// </summary>
        /// <param name="message">The message to log</param>
        public void Success(string message)
        {
            WriteLine(message, ConsoleColor.Green);
        }
    }
}
