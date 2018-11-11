using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;

namespace Repautomator
{
    public static class Config
    {
        public static IConfigurationRoot GenerateConfig(string[] commandArguments)
        {
            var switchMappings = new Dictionary<string, string>
            {
                { "-ReportConfigurationFile", "ReportConfigurationFile"},
                { "-EarliestTime", "EarliestTime"},
                { "-LatestTime", "LatestTime"}
            };

            IConfigurationRoot configCmd = new ConfigurationBuilder()
                .AddCommandLine(commandArguments, switchMappings)
                .Build();

            string pathReportConfig = Path.GetFullPath(configCmd["ReportConfigurationFile"]);
            IConfigurationRoot config = new ConfigurationBuilder()
                .SetBasePath(Path.GetDirectoryName(pathReportConfig))
                .AddJsonFile(Path.GetFileName(pathReportConfig))
                .Build();

            //Convert relative paths to absolute paths
            config["ReportConfiguration:TemplateFile"] = Path.GetFullPath(config["ReportConfiguration:TemplateFile"]);
            config["Outputs:File:Directory"] = Path.GetFullPath(config["Outputs:File:Directory"]);

            //Add default report parameters / custom values
            config["ReportParameters:ReportDateTime"] = DateTime.Now.ToString("s");
            config["ReportParameters:ReportLongDate"] = DateTime.Now.ToString("d MMMM yyyy");
            config["ReportParameters:EarliestTime"] = Tools.ParseInputDate(configCmd["EarliestTime"]).ToString("s");
            config["ReportParameters:LatestTime"] = Tools.ParseInputDate(configCmd["LatestTime"]).ToString("s");


            if (Convert.ToBoolean(config["Inputs:Splunk:Enabled"]))
            {
                string pathSplunkConfig = Path.GetFullPath(config["Inputs:Splunk:ConfigurationFile"]);
                IConfigurationRoot configSplunk = new ConfigurationBuilder()
                    .SetBasePath(Path.GetDirectoryName(pathSplunkConfig))
                    .AddJsonFile(Path.GetFileName(pathSplunkConfig))
                    .Build();
                //Populate config with Splunk server details
                foreach (var item in configSplunk.GetChildren())
                {
                    string configKeyName = String.Format("Inputs:Splunk:Config:{0}", item.Key);
                    config[configKeyName] = item.Value;
                }
                //Substitute the search parameters into the search queries and add the full search text to the global config.
                var splunkQueryCategories = config.GetSection("Inputs:Splunk:Queries").GetChildren();
                foreach (var category in splunkQueryCategories)
                {
                    foreach (var query in category.GetChildren())
                    {
                        string pathQueryFile = Path.GetFullPath(query["FilePath"]);
                        string querySPL = File.ReadAllText(pathQueryFile);

                        foreach (var parameter in config.GetSection("SearchParameters").GetChildren())
                        {
                            querySPL = querySPL.Replace("$" + parameter.Key + "$", parameter.Value);
                        }

                        string configKeyName = String.Format("Inputs:Splunk:Queries:{0}:{1}:Code", category.Key, query.Key);
                        config[configKeyName] = querySPL;
                    }
                }

            }

            //Load email templates from disk
            if (Convert.ToBoolean(config["Outputs:Email:Enabled"]))
            {
                string plaintextTemplateKey = "Outputs:Email:Templates:Body:Plaintext:Template";
                string htmlTemplateKey = "Outputs:Email:Templates:Body:Html:Template";
                config[plaintextTemplateKey] = File.ReadAllText(config[plaintextTemplateKey]);
                config[htmlTemplateKey] = File.ReadAllText(config[htmlTemplateKey]);
            }

            return config;
        }

        public static void WriteConfigSection(IConfigurationSection section, int level)
        {
            var sections = section.GetChildren();
            foreach (var child in sections)
            {
                var indentKey = String.Empty;
                var indentValue = String.Empty;
                var key = child.Key;
                var value = child.Value;
                var path = child.Path;

                if (key.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    value = "**REDACTED**";
                }

                for (int i = 0; i < level - 1; i++)
                {
                    indentKey += "  ";
                    indentValue += "  ";
                }

                if (level > 0)
                {
                    indentKey += "└─";
                    indentValue += "  ";
                }

                if (value == null)
                {
                    Console.WriteLine(String.Format("{0}┬{1}", indentKey, path));
                }
                else
                {
                    Console.WriteLine(String.Format("{0}─{1}", indentKey, path));
                    Console.Write(indentValue);
                    Console.WriteLine(String.Format(" {0}", value));
                }
                WriteConfigSection(child, level + 1);
            }
        }


    }
}
