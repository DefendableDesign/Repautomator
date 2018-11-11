# Repautomator
Repautomator generates beautiful reports, in Word format, from source data that lives in Splunk.

# Setup
1. Compile and publish the app for your target environment
    - `dotnet publish -c release -r win10-x64` 
    - `dotnet publish -c release -r centos-x64`
    - See: https://docs.microsoft.com/en-us/dotnet/core/rid-catalog
1. Update the Splunk API credentials in the `DataSources/splunk.json` configuration file
1. Make sure your output directory exists before running Repautomator as it will not automatically try to create a directory that does not exist.

# Usage
Repautomator supports relative file paths in configuration files. If using relative paths, make sure you run Repautomator.exe from the correct directory.

Usage: `Repautomator.exe -ReportConfigurationFile "ReportConfigs\example.json" -EarliestTime "today-7d" -LatestTime "today"`
This example will generate a sample report and should work with any Splunk setup.

## Repautomator time format
Repautomator uses relative times for running reports, using the following custom format:

`RelativeTimeReference[+/-IntegerModifier]`

### Relative Time References
- `thismonth`:
    - 00:00 on the first of the current month
- `lastmonth`:
    - 00:00 on the first of the previous month
- `nextmonth`:
    - 00:00 on the first of the next month
- `today`:
    - 00:00 today
- `yesterday`:
    - 00:00 yesterday
- `tomorrow`:
    - 00:00 tomorrow

### Time Modifiers
- `d`: Days
- `h`: Hours
- `m`: Minutes
- `M`: Months
- `y`: Years

### Examples
- `thismonth-1Y`
    - 00:00 on the first day of the current month one year ago
- `today+9h`
    - 09:00 today
- `yesterday+9h`
    - 09:00 yesterday