using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Xml;
using static Repautomator.SplunkDataSource;

namespace Repautomator
{
    public class SplunkService
    {
        public string AuthToken { get; private set; }
        public string Host { get; private set; }
        public int Port { get; private set; }
        public int MaxCount { get; private set; }
        public int SearchJobTtl { get; private set; }
        public bool UseSSL { get; private set; }
        public bool ValidateCertificate { get; private set; }
        public enum OutputMode { atom, csv, json, json_cols, json_rows, raw, xml }

        /// <summary>
        /// SplunkService Constructor
        /// </summary>
        /// <param name="host">The hostname of the splunk instance.</param>
        /// <param name="port">The management port for the splunk instance.</param>
        /// <param name="useSSL">True for SSL/HTTPS, false for plain HTTP (not recommended).</param>
        public SplunkService(string host, int port, int maxCount, int searchJobTtl, bool useSSL, bool validateCertificate)
        {
            Host = host;
            Port = port;
            MaxCount = maxCount;
            SearchJobTtl = searchJobTtl;
            UseSSL = useSSL;
            ValidateCertificate = validateCertificate;
        }

        /// <summary>
        /// Send a ServerRequest to the splunk API service.
        /// </summary>
        /// <param name="request">ServerRequest object containing the request details.</param>
        /// <returns>ServerResponse object that contains the details of the response from the server.</returns>
        public ServerResponse Send(ServerRequest request)
        {
            ServerResponse response;
            string url = BuildUrl(request.Path);

            var clientHandler = new HttpClientHandler();
            if (!ValidateCertificate)
            {
                clientHandler.ServerCertificateCustomValidationCallback += (message, certificate2, arg3, arg4) => true;
            }

            using (HttpClient hc = new HttpClient(clientHandler))
            {

                if (AuthToken != null) { hc.DefaultRequestHeaders.Add("Authorization", "Splunk " + AuthToken); }
                try
                {
                    switch (request.Method)
                    {
                        case ServerRequest.HttpMethod.GET:
                            string queryString = request.ArgsQueryString();
                            url = (queryString.Length > 0) ? string.Format("{0}?{1}", url, request.ArgsQueryString()) : url;

                            using (var rm = hc.GetAsync(url).Result)
                            {
                                string responseContent = rm.Content.ReadAsStringAsync().Result;
                                int statusCode = Convert.ToInt32(rm.StatusCode);
                                response = new ServerResponse(responseContent, statusCode);
                            }
                            break;
                        case ServerRequest.HttpMethod.POST:
                            using (var rm = hc.PostAsync(url, new FormUrlEncodedContent(request.Args)).Result)
                            {
                                string responseContent = rm.Content.ReadAsStringAsync().Result;
                                int statusCode = Convert.ToInt32(rm.StatusCode);
                                response = new ServerResponse(responseContent, statusCode);
                            }
                            break;
                        default:
                            throw new NotImplementedException();
                    }
                }
                catch (Exception e)
                {
                    throw new Exception(String.Format("A problem occurred while communicating with the Splunk API.\nCheck that the Splunk API is online and accessible.\n\nThe exception was:\n\n{0}", e.Message));
                }

            }

            return response;
        }

        /// <summary>
        /// Submit a job to Splunk to be scheduled immediately.
        /// </summary>
        /// <param name="job">Job object containing details of the search job.</param>
        /// <returns>Returns the string containing the Splunk job identifier.</returns>
        public string SubmitJob(SplunkDataQuery job)
        {
            string path = "/services/search/jobs";
            ServerRequest request = new ServerRequest(path, ServerRequest.HttpMethod.POST);
            request.Args.Add(new KeyValuePair<string, string>("search", job.Value));
            request.Args.Add(new KeyValuePair<string, string>("earliest_time", string.Format("{0}.000+00:00", job.EarliestTime.ToUniversalTime().ToString("s"))));
            request.Args.Add(new KeyValuePair<string, string>("latest_time", string.Format("{0}.000+00:00", job.LatestTime.ToUniversalTime().ToString("s"))));
            request.Args.Add(new KeyValuePair<string, string>("max_count", MaxCount.ToString()));
            request.Args.Add(new KeyValuePair<string, string>("timeout", SearchJobTtl.ToString()));

            ServerResponse response = this.Send(request);

            var doc = new XmlDocument();
            doc.LoadXml(response.Content);
            string sid;
            try
            {
                sid = doc.SelectSingleNode("/response/sid").InnerText;
            }
            catch (Exception)
            {
                throw new Exception(String.Format("Something went wrong while submitting the search to Splunk. The Splunk API returned:\n{0}", response.Content));
            }
            
            return sid;
        }

        /// <summary>
        /// Returns the current dispatch state of supplied job.
        /// </summary>
        /// <param name="job">Job object containing details of the search job.</param>
        /// <returns>The 'dispatchStatus'.</returns>
        public SplunkJobStatus GetJobStatus(string searchId)
        {
            string path = string.Format("{0}/{1}", "/services/search/jobs", searchId);
            ServerRequest request = new ServerRequest(path, ServerRequest.HttpMethod.GET);
            ServerResponse response = this.Send(request);

            var doc = new XmlDocument();
            doc.LoadXml(response.Content);

            //Create XmlNamespaceManager so that we can query s: prefixed nodes.
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            //Add Splunk and OpenSearch namespaces
            nsmgr.AddNamespace("s", "http://dev.splunk.com/ns/rest");
            nsmgr.AddNamespace("opensearch", "http://a9.com/-/spec/opensearch/1.1/");

            var dispatchStatus = doc.SelectSingleNode("//s:dict/s:key[@name='dispatchState']", nsmgr).InnerText;

            switch (dispatchStatus)
            {
                case "QUEUED":
                    return SplunkJobStatus.QUEUED;
                case "PARSING":
                    return SplunkJobStatus.PARSING;
                case "RUNNING":
                    return SplunkJobStatus.RUNNING;
                case "PAUSED":
                    return SplunkJobStatus.PAUSED;
                case "FINALIZING":
                    return SplunkJobStatus.FINALIZING;
                case "FAILED":
                    return SplunkJobStatus.FAILED;
                case "DONE":
                    return SplunkJobStatus.DONE;
                default:
                    throw new KeyNotFoundException(String.Format("Unknown job status: {0}", dispatchStatus));
            }
        }

        public string GetJobResults(SplunkDataQuery job, OutputMode mode)
        {
            string path = string.Format("{0}/{1}/{2}", "/services/search/jobs", job.RemoteId, "results");
            ServerRequest request = new ServerRequest(path, ServerRequest.HttpMethod.GET);

            //Set count to 0 to get all rows
            request.Args.Add(new KeyValuePair<string, string>("count", "0"));
            request.Args.Add(new KeyValuePair<string, string>("output_mode", mode.ToString()));

            ServerResponse response = this.Send(request);
            if (response.Status == 204)
            {
                return null;
            }
            else
            {
                return response.Content;
            }
        }

        /// <summary>
        /// Logs in to Splunk and sets the current authorization token.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <param name="password">The password.</param>
        public void Login(string username, string password)
        {
            ServerRequest request = new ServerRequest("/services/auth/login", ServerRequest.HttpMethod.POST);
            request.Args.Add(new KeyValuePair<string, string>("username", username));
            request.Args.Add(new KeyValuePair<string, string>("password", password));

            ServerResponse loginResponse = this.Send(request);

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(loginResponse.Content);

            string authToken = doc.SelectSingleNode("/response/sessionKey").InnerText;
            this.AuthToken = authToken;
        }

        /// <summary>
        /// This method returns a full URL to the Splunk instance for a given path.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <returns>The URL.</returns>
        private string BuildUrl(string path)
        {
            string scope = UseSSL ? "https" : "http";
            return string.Format("{0}://{1}:{2}{3}", scope, Host, Port, path);
        }
    }
}
