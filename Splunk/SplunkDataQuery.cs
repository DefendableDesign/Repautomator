using System;
using static Repautomator.SplunkDataSource;

namespace Repautomator
{
    /// <summary>
    /// A class for managing Splunk search queries.
    /// </summary>
    public class SplunkDataQuery : IDataQuery
    {
        public string Key { get; private set; }
        public string Value { get; private set; }
        public DateTime EarliestTime { get; private set; }
        public DateTime LatestTime { get; private set; }
        
        public string RemoteId { get; private set; }
        public QueryStatus Status { get { return getStatus(); } }
        public string Result { get { return getResults(); } }


        private int MaxCount { get; set; }
        private string searchResults { get; set; }
        private SplunkJobStatus SplunkStatus { get { return getSplunkStatus(); } }
        private SplunkJobStatus LastKnownSplunkStatus { get; set; }
        private SplunkService Service { get; set; }

        /// <summary>
        /// Query Constructor
        /// </summary>
        /// <param name="key">The name of the Search/Job. It does not need to be globally unique.</param>
        /// <param name="value">The Splunk Processing Language (SPL) for the search query.</param>
        /// <param name="service">The SplunkService object to provide access to the Splunk API.</param>
        /// <param name="earliestTime">
        /// The earliest event time
        /// </param>
        /// <param name="latestTime">
        /// The latest event time
        /// </param>
        /// <param name="maxEventCount">The maximum number of events to retrieve for each search</param>
        public SplunkDataQuery(string key, string value, SplunkService service, DateTime earliestTime, DateTime latestTime)
        {
            searchResults = null;

            Key = key;
            Value = value;
            Service = service;
            EarliestTime = earliestTime;
            LatestTime = latestTime;

            RemoteId = Service.SubmitJob(this);
        }

        private QueryStatus getStatus()
        {
            switch (SplunkStatus)
            {
                case SplunkJobStatus.QUEUED:
                    return QueryStatus.WAITING;
                case SplunkJobStatus.PARSING:
                    return QueryStatus.WAITING;
                case SplunkJobStatus.RUNNING:
                    return QueryStatus.RUNNING;
                case SplunkJobStatus.PAUSED:
                    return QueryStatus.WAITING;
                case SplunkJobStatus.FINALIZING:
                    return QueryStatus.RUNNING;
                case SplunkJobStatus.FAILED:
                    return QueryStatus.ERROR;
                case SplunkJobStatus.DONE:
                    return QueryStatus.COMPLETED;
                default:
                    return QueryStatus.ERROR;
            }
        }

        /// <summary>
        /// Returns the status of the search job from the Splunk API.
        /// </summary>
        /// <returns>Boolean</returns>
        public bool IsComplete()
        {
            if (SplunkStatus == SplunkJobStatus.DONE) return true;
            if (SplunkStatus == SplunkJobStatus.FAILED) throw new Exception(String.Format("Splunk job ({0}:{1}) failed.", RemoteId, Key));
            else return false;
        }

        /// <summary>
        /// Gets the status of any running jobs from the Splunk API.
        /// </summary>
        /// <returns>Returns the SplunkJobStatus of the query's search job.</returns>
        private SplunkJobStatus getSplunkStatus()
        {
            switch (LastKnownSplunkStatus)
            {
                case SplunkJobStatus.FAILED:
                    return SplunkJobStatus.FAILED;
                case SplunkJobStatus.DONE:
                    return SplunkJobStatus.DONE;
                default:
                    LastKnownSplunkStatus = Service.GetJobStatus(RemoteId);
                    return LastKnownSplunkStatus;
            }
        }

        /// <summary>
        /// Gets the results of the search from the Splunk API.
        /// </summary>
        /// <returns>Returns the searchResults in JSON format.</returns>
        private string getResults()
        {
            if (SplunkStatus == SplunkJobStatus.DONE && searchResults == null)
            {
                searchResults = Service.GetJobResults(this, SplunkService.OutputMode.json_rows);
            }

            return searchResults;
        }


    }
}
