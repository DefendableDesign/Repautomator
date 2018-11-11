using System;

namespace Repautomator
{
    public interface IDataQuery
    {
        string Key { get; }
        string Value { get; }
        DateTime EarliestTime { get; }
        DateTime LatestTime { get; }

        string RemoteId { get; }
        QueryStatus Status { get; }
        string Result { get; }

        bool IsComplete();        
    }

    public enum QueryStatus
    {
        WAITING,
        RUNNING,
        COMPLETED,
        ERROR
    }
}
