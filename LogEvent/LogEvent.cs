using System;
using System.IO;

namespace ProValAPI
{
    class LogEvent
    {        
        static string FOLDER = null;
        
        // Example #1: Write an array of strings to a file.
        // Create a string array that consists of three lines.

        public static void Log(string logTest)
        {
            if (FOLDER == null) {
                FOLDER = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                FOLDER = FOLDER + "\\WinTech\\PVAPI\\Log\\";
                var d = Directory.CreateDirectory(FOLDER);
            }
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            var today = DateTime.Today.ToString("yyyyMMdd");
            string logfile = FOLDER + @"log_" + today + ".txt";
            using (StreamWriter file =
                new StreamWriter(logfile, true))
            {
                file.WriteLine(timestamp + ":  " + logTest);
            }
        }
    }
}
