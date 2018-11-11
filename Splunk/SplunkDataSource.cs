using System;

namespace Repautomator
{
    public class SplunkDataSource : IDataSource
    {
        public SplunkService Service { get; private set; }
        
        public SplunkDataSource(string hostname, int port, string username, string password, int maxCount, int searchJobTtl, bool useTls = true, bool validateCertificate = true)
        {
            Service = new SplunkService(hostname, port, maxCount, searchJobTtl, useTls, validateCertificate);
            Service.Login(username, password);
        }

        /// <summary>
        /// Runs a new query against the current DataSource's Splunk instance
        /// </summary>
        /// <param name="key">The name of the Search/Job. It does not need to be globally unique.</param>
        /// <param name="value">The Splunk Processing Language (SPL) for the search query.</param>
        /// <param name="earliestTime">
        /// The earliest event time in Splunk time format. eg. -1d@d or %m/%d/%Y:%H:%M:%S
        /// See http://docs.splunk.com/Documentation/Splunk/6.5.2/Search/Specifytimemodifiersinyoursearch
        /// </param>
        /// <param name="latestTime">
        /// The latest event time in Splunk time format. eg. @d or %m/%d/%Y:%H:%M:%S
        /// See http://docs.splunk.com/Documentation/Splunk/6.5.2/Search/Specifytimemodifiersinyoursearch
        /// </param>
        /// <returns>SplunkQuery object to manage the Splunk search query job and results. </returns>
        public IDataQuery Query(string key, string value, DateTime earliestTime, DateTime latestTime)
        {
            var result = new SplunkDataQuery(key, value, Service, earliestTime, latestTime);
            return result;
        }

        public enum SplunkJobStatus
        {
            QUEUED,
            PARSING,
            RUNNING,
            PAUSED,
            FINALIZING,
            FAILED,
            DONE
        }
    }
}
