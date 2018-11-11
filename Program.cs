using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using static Repautomator.Tools;
using static Repautomator.Config;
//using Console = System.Console;


namespace Repautomator
{
    public class Program
    {
        private static int CurrentStep { get; set; }
        private static List<IDataQuery> Queries { get; set; }
        private static IConfigurationRoot Config { get; set; }
        private static SplunkDataSource Splunk { get; set; }

        public static void Main(string[] args)
        {
            DateTime startTime = DateTime.Now;
            Config = GenerateConfig(args);
            Queries = new List<IDataQuery>();


            SetupConsole();
            ProcessInputs();
            WaitForInputs();
            ProcessResults();
            BuildReport();

            DebugConfig();

            TimeSpan duration = DateTime.Now - startTime;
            Console.WriteLine(String.Format("\nExecution completed in {0}h {1}m {2}s.", duration.Hours, duration.Minutes, duration.Seconds));


        }

        /// <summary>
        /// Configures the Console encoding and colours and writes the Program banner and configuration.
        /// </summary>
        private static void SetupConsole()
        {
            string header = @"8888888b.                                     888                                   888                    
888   Y88b                                    888                                   888                    
888    888                                    888                                   888                    
888   d88P .d88b.  88888b.   8888b.  888  888 888888 .d88b.  88888b.d88b.   8888b.  888888 .d88b.  888d888 
8888888P"" d8P  Y8b 888 ""88b     ""88b 888  888 888   d88""""88b 888 ""888 ""88b     ""88b 888   d88""""88b 888P""   
888 T88b  88888888 888  888 .d888888 888  888 888   888  888 888  888  888 .d888888 888   888  888 888     
888  T88b Y8b.     888 d88P 888  888 Y88b 888 Y88b. Y88..88P 888  888  888 888  888 Y88b. Y88..88P 888     
888   T88b ""Y8888  88888P""  ""Y888888  ""Y88888  ""Y888 ""Y88P""  888  888  888 ""Y888888  ""Y888 ""Y88P""  888     
                   888                                                                                     
                   888                                                                                     
                   888                                                                                     ";

            Console.WriteLine(header + "\n");

            List<IConfigurationSection> parameters = new List<IConfigurationSection>();
            parameters.AddRange(Config.GetSection("ReportParameters").GetChildren());
            parameters.AddRange(Config.GetSection("Paths").GetChildren());
            parameters.WriteToConsole();
            Console.WriteLine();
        }

        /// <summary>
        /// Configures the required IDataSource(s) and launches the queries.
        /// </summary>
        private static void ProcessInputs()
        {
            WriteStep("Processing Inputs");
            if (Convert.ToBoolean(Config["Inputs:Splunk:Enabled"])) ProcessSplunkInputs();
            WriteStepComplete();
        }

        /// <summary>
        /// Processing loop to wait for running IDataQuery(s) to finish.
        /// </summary>
        private static void WaitForInputs()
        {
            WriteStep("Waiting for Inputs");
            var countTotal = Queries.Count;
            var countComplete = Queries.Where(e => e.IsComplete()).Count();
            var top = (System.Console.IsOutputRedirected || System.Console.IsErrorRedirected) ? 0 : System.Console.CursorTop;
            while (countComplete < countTotal)
            {
                int pctComplete = Convert.ToInt32((Convert.ToDouble(countComplete) / countTotal) * 100);
                if (!System.Console.IsOutputRedirected || !System.Console.IsErrorRedirected) ProgressBar(pctComplete, top, 0);
                Thread.Sleep(1000);
                countComplete = countComplete = Queries.Where(e => e.IsComplete()).Count();
            }
            if (!System.Console.IsOutputRedirected || !System.Console.IsErrorRedirected)
            {
                ProgressBar(100, top, 0);
                ProgressBar(null, top, 0);
            }
            WriteStepComplete();
        }

        /// <summary>
        /// Retrieves the completed IDataQuery(s) results and injects them into the Program Config.
        /// </summary>
        private static void ProcessResults()
        {
            WriteStep("Processing Query Results");
            foreach (var query in Queries)
            {
                var configKey = String.Format("{0}:Result", query.Key);
                Config[configKey] = query.Result;
            }
            WriteStepComplete();
        }

        /// <summary>
        /// Creates the WordprocessingDocument report and outputs it according to the configured outputs.
        /// </summary>
        private static void BuildReport()
        {
            WriteStep("Building Report");
            MemoryStream fileContents = ReportBuilder.BuildReport(Config);
            string fileName = null;

            if (Convert.ToBoolean(Config["Outputs:File:Enabled"]))
            {
                fileName = MakeValidFileName(ProcessTemplate(Config, "Outputs:File:FileName"));
                var outputPath = Path.Combine(Config["Outputs:File:Directory"], fileName);
                WriteFile(fileContents, outputPath);
                Thread.Sleep(5000);
            }

            if (Convert.ToBoolean(Config["Outputs:Email:Enabled"]))
            {
                fileContents.Position = 0;
                fileName = MakeValidFileName(ProcessTemplate(Config, "Outputs:Email:Templates:AttachmentFileName"));
                EmailReport(Config, fileContents, fileName);
            }

            WriteStepComplete();
        }

        /// <summary>
        /// Configures the SplunkDataSource and launches the SplunkDataQuery(s).
        /// </summary>
        private static void ProcessSplunkInputs()
        {
            Splunk = new SplunkDataSource(
               Config["Inputs:Splunk:Config:Hostname"],
               Convert.ToInt32(Config["Inputs:Splunk:Config:Port"]),
               Config["Inputs:Splunk:Config:Username"],
               Config["Inputs:Splunk:Config:Password"],
               Convert.ToInt32(Config["Inputs:Splunk:Config:MaxCount"]),
               Convert.ToInt32(Config["Inputs:Splunk:Config:SearchJobTtl"]),
               Convert.ToBoolean(Config["Inputs:Splunk:Config:UseTls"]),
               Convert.ToBoolean(Config["Inputs:Splunk:Config:ValidateCertificate"])
               );

            foreach (var category in Config.GetSection("Inputs:Splunk:Queries").GetChildren())
            {
                Console.WriteLine(String.Format("Starting {0} queries:", category.Key));
                foreach (var query in category.GetChildren())
                {
                    var sq = Splunk.Query(query.Path, query["Code"], Convert.ToDateTime(Config["ReportParameters:EarliestTime"]), Convert.ToDateTime(Config["ReportParameters:LatestTime"]));
                    Queries.Add(sq);
                    Console.WriteLine(String.Format("\t{0} query submitted", query.Key));
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// Dumps the running Config to screen.
        /// </summary>
        private static void DebugConfig()
        {
            WriteStep("Dumping in-memory configuration");
            if (Convert.ToBoolean(Config["Debug"]))
            {
                foreach (var section in Config.GetChildren())
                {
                    WriteConfigSection(section, 0);
                }
            }
        }

        /// <summary>
        /// Writes the current step to screen.
        /// </summary>
        /// <param name="message">Description of the current step.</param>
        public static void WriteStep(string message)
        {
            CurrentStep++;

            if (System.Console.IsOutputRedirected || System.Console.IsErrorRedirected)
            {
                Console.WriteLine(String.Format("{0}: Step {1}: {2} - Started", DateTime.Now.ToString("s"), CurrentStep, message));
            }
            else
            {
                Console.Write(DateTime.Now.ToString("s") + ": ");
                Console.WriteLine(String.Format("Step {0}: {1}\n", CurrentStep, message));
            }
            
        }

        /// <summary>
        /// Writes the completion message for the current step to screen.
        /// </summary>
        public static void WriteStepComplete()
        {
            if (System.Console.IsOutputRedirected || System.Console.IsErrorRedirected)
            {
                Console.WriteLine(String.Format("{0}: Step {1}: Complete", DateTime.Now.ToString("s"), CurrentStep));
            }
            else
            {
                Console.Write("\t√ ");
                Console.WriteLine("Step complete.\n");
            }

        }
    }
}
