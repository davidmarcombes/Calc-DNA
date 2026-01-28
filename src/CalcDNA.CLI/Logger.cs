using System;
using System.Collections.Generic;
using System.Text;

namespace CalcDNA.CLI
{
    public class Logger
    {
        // Note: simplistic logger for console output, function no static on purpose to allow future enhancements

        public void Info(string message)
        {
            Console.WriteLine(message);
        }

        public void Error(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            try
            {
                Console.Error.WriteLine(message);
            }
            finally
            {
                Console.ResetColor();
            }
        }

        public void Warning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            try
            {
                Console.Error.WriteLine(message);
            }
            finally
            {
                Console.ResetColor();
            }
        }

        public void Debug(string message, bool verbose)
        {
            if (verbose)
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                try
                {
                    Console.WriteLine(message);
                }
                finally
                {
                    Console.ResetColor();
                }
            }
        }

        public void Success(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            try
            {
                Console.WriteLine(message);
            }
            finally
            {
                Console.ResetColor();
            }
        }

    }
}
