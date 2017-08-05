using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace FogBugzImporter
{
    class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            string url = "http://localhost:8080";
			string name = "max";
			string password = "pass";
			string ticketsFile = "/Users/Max/Desktop/Import.csv";
#else
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: URL Username Password TicketsFile.\n");
                Console.WriteLine("Example: http://localhost/fb/api.asp \"My Name\" \"Very Secret Password\" tickets.xls.\n");
                return;
            }

            string url = args[0];
            string name = args[1];
            string password = args[2];
            string ticketsFile = args[3];
#endif

			if (!File.Exists(ticketsFile))
            {
                Console.WriteLine("Tickets file not found.\n");
                return;
            }

            string sourceDir = Path.GetDirectoryName(ticketsFile);
            string mediaDir = Path.Combine(sourceDir, "media");

            if (!Directory.Exists(mediaDir))
            {
                Console.WriteLine("Directory with attachments data not found.\n");
                //return;
            }

            Importer importer = new Importer(url, name, password, ticketsFile, mediaDir);
            //importer.Connect();
            importer.Import();
        }
    }
}
