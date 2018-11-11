using System;

namespace Repautomator
{
    interface IDataSource
    {
        IDataQuery Query(string key, string value, DateTime earliestTime, DateTime latestTime);        
    }
}
