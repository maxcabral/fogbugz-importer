using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Mono.Options;

namespace FogBugzImporter
{
    class Program
    {
        static string url = null;
		static string emailAddress = null;
		static string password = null;
		static List<string> ticketsFiles = new List<string>();
		static string token = null;
		static bool force = false;
		static bool verbose = true;
		static bool shouldShowHelp = false;

        static void Main(string[] args)
        {
            //ProcessOptions(args);

            var driver = new FogBugzDriver(url, token);
            try
            {
                driver.ConnectAsync().Wait();
            } catch (Exception ex){
                Console.WriteLine("Unable to connect to FogBugz");
                if (verbose){
                    Console.WriteLine(ex);
                }
                Environment.Exit(1);
            }

            var tasks = new List<Task>(ticketsFiles.Count());
            foreach (var ticketsFile in ticketsFiles)
            {
                Console.WriteLine($"Processing {ticketsFile}");
				string sourceDir = Path.GetDirectoryName(ticketsFile);
				string mediaDir = Path.Combine(sourceDir, "media");

                var importer = new Importer(driver, ticketsFile, mediaDir);
                tasks.Add(importer.ImportAsync());
            }
            Task.WaitAll(tasks.ToArray());
        }

        static void ProcessOptions(string[] args)
        {
			var options = new OptionSet {
				{ "e|email=", "Your FogBugz Login Email Address", e => emailAddress = e },
				{ "p|password=", "Your FogBugz Login Password", p => password = p },
				{ "t|token=", "Your FogBugz API Token (recommended)", t => token = t },
				{ "u|api-url=", "Your FogBugz API Url, i.e. https://example.fogbugz.com/api.asp", u => url = u },
				{ "i|import", "A path to a file to import. Multiple files may be specified.", f => ticketsFiles.Add(f) },
				{ "f|force", "Ignore most errors", f => force = f != null },
				{ "v|verbose", "Increase debug message verbosity", v => verbose = v != null },
				{ "h|help", "Show this message and exit", h => shouldShowHelp = h != null },
			};

			try
			{
				options.Parse(args);

				if (shouldShowHelp)
				{
					Console.Write("Usage: FogBugzImporter.exe [OPTIONS]\nImports CSVs into FogBugz\n\n");

					// output the options
					Console.WriteLine("Options:");
					options.WriteOptionDescriptions(Console.Out);
					Environment.Exit(0);
				}

				//Make sure we don't see nothing but empty files
				ticketsFiles = ticketsFiles.Where(ticketFile => !String.IsNullOrWhiteSpace(ticketFile)).ToList();

				if (ticketsFiles.Count == 0 || String.IsNullOrWhiteSpace(url))
				{
					throw new OptionException("At least one import is required", "i|import");
				}

				if (String.IsNullOrWhiteSpace(url))
				{
					throw new OptionException("api-url is required", "u|api-url");
				}

				if (String.IsNullOrWhiteSpace(token) && (String.IsNullOrWhiteSpace(emailAddress) || String.IsNullOrWhiteSpace(password)))
				{
					throw new OptionException("api-url or email and password are required", "u|api-url or e|email and p|password");
				}
			}
			catch (OptionException e)
			{
				// output some error message
				Console.Write("fogbugz-importer: ");
				Console.WriteLine(e.Message);
				Console.WriteLine("Try `fogbugz-importer --help' for more information.");
                Environment.Exit(1);
			}

			var lockedFiles = new List<String>();
			var missingFiles = new List<String>();
			ticketsFiles = ticketsFiles.Where((ticketFile) =>
			{
				if (!File.Exists(ticketFile))
				{
					missingFiles.Add(ticketFile);
					Console.WriteLine("Tickets file not found.\n");
					return false;
				}

				try
				{
					File.Open(ticketFile, FileMode.Open, FileAccess.Read).Dispose();
				}
				catch (IOException)
				{
					lockedFiles.Add(ticketFile);
					Console.WriteLine("\n");
					return false;
				}

				return true;
			}).ToList();

			if (!force && (lockedFiles.Count != 0 || missingFiles.Count != 0))
			{
				Console.Write("Exiting due to errors with the following files:\n\n");

				if (lockedFiles.Count != 0)
				{
					Console.Write("These files can not be read. Another program may have opened them:\n\n{0}\n\n", String.Join("\n", lockedFiles));
				}

				if (missingFiles.Count != 0)
				{
					Console.Write("These file can not be found:\n\n{0}\n\n", String.Join("\n", missingFiles));
				}

				Environment.Exit(1);
			}
        }
    }
}
