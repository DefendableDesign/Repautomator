using System.Net;
using System.Collections.Generic;
using System.Text;

namespace Repautomator
{
    public class ServerRequest
    {
        public enum HttpMethod { GET, POST, DELETE }
        private string path;
        private HttpMethod method;
        private List<KeyValuePair<string, string>> args;

        public string Path { get { return path; } }
        public HttpMethod Method { get { return method; } }
        public List<KeyValuePair<string, string>> Args { get { return args; } }

        /// <summary>
        /// ServerRequest Constructor
        /// </summary>
        /// <param name="path">REST URI path of the Splunk API method.</param>
        /// <param name="method">HTTP Method (GET, POST, DELETE).</param>
        public ServerRequest(string path, HttpMethod method)
        {
            this.path = path;
            this.method = method;
            this.args = new List<KeyValuePair<string, string>>();
        }

        /// <summary>
        /// ServerRequest Constructor
        /// </summary>
        /// <param name="path">REST URI path of the Splunk API method.</param>
        /// <param name="method">HTTP Method (GET, POST, DELETE).</param>
        /// <param name="args">List of KeyValuePairs containing the arguments to send to the API service.</param>
        public ServerRequest(string path, HttpMethod method, List<KeyValuePair<string, string>> args)
        {
            this.path = path;
            this.method = method;
            this.args = args;
        }

        /// <summary>
        /// Returns the current request arguments in query string format.
        /// </summary>
        /// <returns>The query string.</returns>
        public string ArgsQueryString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var kvp in Args)
            {
                string key = WebUtility.UrlEncode(kvp.Key);
                string value = WebUtility.UrlEncode(kvp.Value);
                if (sb.Length != 0) { sb.Append("&"); }
                sb.Append(string.Format("{0}={1}", key, value));
            }
            return sb.ToString();
        }
    }
}
