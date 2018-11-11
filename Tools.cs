using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using MailKit.Net.Smtp;
using MimeKit;
using MimeKit.Utils;

namespace Repautomator
{
    public static class Tools
    {
        enum DtModifiers { Add, Subtract };
        public static DateTime ParseInputDate(string dt)
        {
            DtModifiers modifier = DtModifiers.Add;
            DateTime now = DateTime.Now;
            DateTime realDate;
            int modNumber = 0;
            string modValue = "";

            if (dt.Contains('-')) { modifier = DtModifiers.Subtract; }
            else if (dt.Contains('+')) { modifier = DtModifiers.Add; }

            string[] mods = dt.Split('+', '-');

            if (mods.Length > 2) { throw new ArgumentException("The provided datetime string is not recognised"); }

            try
            {
                if (mods.Length == 2)
                {
                    Match match = Regex.Match(mods[1], @"(\d+)(\w+)");
                    modNumber = Convert.ToInt32(match.Groups[1].Value);
                    modValue = match.Groups[2].Value;
                }
            }
            catch (Exception)
            {
                throw new ArgumentException("The provided datetime string is not recognised");
            }

            switch (mods[0])
            {
                case "thismonth":
                    realDate = new DateTime(now.Year, now.Month, 1, 0, 0, 0);
                    break;
                case "lastmonth":
                    realDate = new DateTime(now.Year, now.Month, 1, 0, 0, 0).AddMonths(-1);
                    break;
                case "nextmonth":
                    realDate = new DateTime(now.Year, now.Month, 1, 0, 0, 0).AddMonths(1);
                    break;
                case "today":
                    realDate = new DateTime(now.Year, now.Month, now.Day, 0, 0, 0);
                    break;
                case "yesterday":
                    realDate = new DateTime(now.Year, now.Month, now.Day, 0, 0, 0).AddDays(-1);
                    break;
                case "tomorrow":
                    realDate = new DateTime(now.Year, now.Month, now.Day, 0, 0, 0).AddDays(1);
                    break;
                default:
                    realDate = now;
                    break;
            }

            //If we're subtracting, convert our positive modNumber to the equivalent negative.
            if (modifier == DtModifiers.Subtract) { modNumber = modNumber * -1; }

            switch (modValue)
            {
                case "d":
                    realDate = realDate.AddDays(modNumber);
                    break;
                case "h":
                    realDate = realDate.AddHours(modNumber);
                    break;
                case "m":
                    realDate = realDate.AddMinutes(modNumber);
                    break;
                case "M":
                    realDate = realDate.AddMonths(modNumber);
                    break;
                case "y":
                    realDate = realDate.AddYears(modNumber);
                    break;
                default:
                    break;
            }

            return realDate;
        }

        public static void ProgressBar(int? pct, int top, int left)
        {
            int width = Console.WindowWidth - 7;
            int completeChars = Convert.ToInt32(pct) * width / 100;

            if (pct < 0 || pct > 100) throw new IndexOutOfRangeException("Percent must be between 0 and 100");

            Console.SetCursorPosition(left, top);

            Console.Write(String.Format(" {0, 3}% ", pct));
            for (int i = 0; i < completeChars; i++) Console.Write("█");
            for (int i = 0 / 2; i < width - completeChars; i++) Console.Write(" ");
            Console.Write(" ");

            if (pct == null)
            {
                Console.SetCursorPosition(left, top);
                for (int i = 0; i < Console.WindowWidth; i++)
                {
                    Console.Write(" ");
                }
                Console.SetCursorPosition(0, top);
            }
        }

        public static void WriteFile(MemoryStream inputStream, string outputPath)
        {
            inputStream.Position = 0;
            using (FileStream fs = new FileStream(outputPath, FileMode.Create))
            {
                inputStream.WriteTo(fs);
            }
        }

        public static string MakeValidFileName(string name)
        {
            string result = name;
            foreach (var item in System.IO.Path.GetInvalidFileNameChars()) { result = result.Replace(item, '_'); }
            return result;
        }

        public static List<string> ParseEmailAddresses(string addresses)
        {
            List<string> result = new List<string>();
            foreach (var address in addresses.Split(';')) { result.Add(address.Trim()); }
            return result;
        }

        public static string SplitCamelCase(string input)
        {
            return System.Text.RegularExpressions.Regex.Replace(input, "([A-Z])", " $1", System.Text.RegularExpressions.RegexOptions.Compiled).Trim();
        }

        public static void WriteToConsole(this List<IConfigurationSection> parameters)
        {
            int colWidth = SplitCamelCase(parameters.Aggregate((max, cur) => SplitCamelCase(max.Key).Length > SplitCamelCase(cur.Key).Length ? max : cur).Key).Length;

            Console.Write(String.Format("{0}: ", "Execution Time".PadRight(colWidth)));
            Console.WriteLine(DateTime.Now.ToString("s"));

            foreach (var parameter in parameters)
            {
                if (parameter.Value != null)
                {
                    Console.Write(String.Format("{0}: ", SplitCamelCase(parameter.Key).PadRight(colWidth)));
                    Console.WriteLine(parameter.Value);
                }
            }
        }

        /// <summary>
        /// Combines a template string with values contained within the 
        /// </summary>
        /// <param name="config">
        /// An IConfigurationSection containing a Template value (in String.Format()) and a Values array containing the names of other Config key paths.</param>
        /// <param name="pathToTemplate">
        /// </param>
        /// <returns></returns>
        public static string ProcessTemplate(IConfigurationRoot config, string pathToTemplateParent)
        {
            string templatePath = String.Format("{0}:Template", pathToTemplateParent);
            string valuesPath = String.Format("{0}:Values", pathToTemplateParent);
            string output;
            string template = config[templatePath];
            string[] valueKeys = config.GetSection(valuesPath).GetChildren().Select(x => x.Value).ToArray();
            int numParams = config.GetSection(valuesPath).GetChildren().Count();
            string[] parameters = new string[numParams];

            for (int i = 0; i < numParams; i++)
            {
                parameters[i] = config[valueKeys[i]];
            }

            output = String.Format(template, parameters);
            return output;
        }

        public static void EmailReport(IConfigurationRoot config, MemoryStream attachmentStream, string attachmentName)
        {
            MimeMessage mail = new MimeMessage();
            attachmentStream.Seek(0, SeekOrigin.Begin);

            string server = config["Outputs:Email:SmtpServer:Host"];
            int port = Convert.ToInt32(config["Outputs:Email:SmtpServer:Port"]);
            bool useTls = Convert.ToBoolean(config["Outputs:Email:SmtpServer:UseTls"]);

            string username = config["Outputs:Email:SmtpServer:Credentials:Username"];
            string password = config["Outputs:Email:SmtpServer:Credentials:Password"];
            username = (username == String.Empty) ? null : username;
            username = (password == String.Empty) ? null : password;

            MailboxAddress fromAddress = null;
            List<MailboxAddress> toAddresses = new List<MailboxAddress>();
            List<MailboxAddress> bccAddresses = new List<MailboxAddress>();

            //Parse From Address
            if (config["Outputs:Email:From:Address"] != null)
            {
                if (config["Outputs:Email:From:Name"] != null)
                {
                    fromAddress = new MailboxAddress(config["Outputs:Email:From:Name"], config["Outputs:Email:From:Address"]);
                }
                else
                {
                    fromAddress = new MailboxAddress(config["Outputs:Email:From:Address"]);
                }
            }

            //Parse To Address(es)
            foreach (var item in config.GetSection("Outputs:Email:To").GetChildren())
            {
                if (item["Address"] != null)
                {
                    if (item["Name"] != null)
                    {
                        toAddresses.Add(new MailboxAddress(item["Name"], item["Address"]));
                    }
                    else
                    {
                        toAddresses.Add(new MailboxAddress(item["Address"]));
                    }
                }
            }

            //Parse BCC Address(es)
            foreach (var item in config.GetSection("Outputs:Email:Bcc").GetChildren())
            {
                if (item["Address"] != null)
                {
                    if (item["Name"] != null)
                    {
                        bccAddresses.Add(new MailboxAddress(item["Name"], item["Address"]));
                    }
                    else
                    {
                        bccAddresses.Add(new MailboxAddress(item["Address"]));
                    }
                }
            }

            //Must have at least one To address or BCC address to send an email.
            if (toAddresses.Count == 0 && bccAddresses.Count == 0) throw new System.ArgumentNullException("Outputs:Email:To / Outputs:Email:Bcc - No addresses found.");

            mail.From.Add(fromAddress);
            if (toAddresses.Count > 0) mail.To.AddRange(toAddresses);
            if (bccAddresses.Count > 0) mail.Bcc.AddRange(bccAddresses);

            mail.Subject = ProcessTemplate(config, "Outputs:Email:Templates:Subject");

            var builder = new BodyBuilder();

            // Load Linked Resources
            foreach (var resource in config.GetSection("Outputs:Email:Templates:Body:Html:LinkedResources").GetChildren())
            {
                var image = builder.LinkedResources.Add(Path.GetFullPath(resource["FilePath"]));
                image.ContentId = MimeUtils.GenerateMessageId();
                config[resource.Path + ":Cid"] = image.ContentId;
            }
            
            // Set the plain-text version of the message text
            builder.TextBody = ProcessTemplate(config, "Outputs:Email:Templates:Body:Plaintext");

            // Set the html version of the message text
            builder.HtmlBody = ProcessTemplate(config, "Outputs:Email:Templates:Body:Html");

            //Attach the report
            ContentType docxContentType = new ContentType("application", "vnd.openxmlformats-officedocument.wordprocessingml.document");
            var attachment = builder.Attachments.Add(attachmentName, attachmentStream, docxContentType);
            attachment.ContentDisposition.CreationDate = DateTime.Now;
            attachment.ContentDisposition.ModificationDate = DateTime.Now;

            // Now we just need to set the message body and we're done
            mail.Body = builder.ToMessageBody();

            using (var client = new SmtpClient())
            {
                if (useTls)
                {
                    client.Connect(server, port, MailKit.Security.SecureSocketOptions.StartTls);
                }
                else
                {
                    client.Connect(server, port, MailKit.Security.SecureSocketOptions.None);
                }

                if (client.AuthenticationMechanisms.Contains("XOAUTH2"))
                {
                    client.AuthenticationMechanisms.Remove("XOAUTH2");
                }

                if (username != null && password != null)
                {
                    client.Authenticate(username, password);
                }

                client.Send(mail);
                client.Disconnect(true);
            }

        }
    }
}
