# Cisco Unified Intelligence Center Statistics
This tiny Python library provides a few functions for importing and interpreting call data from the Cisco Unified Intelligence Center (UIC). Call data exported from UIC are saved as large .xlsx files detailing (at its most granular) half an hour's worth of call data. 
For call center managers, it may be relevant to visualize data and check statistics on what months, days of the week and times of day are busiest to adjust personnel allocation accordingly.

## Requirements
- Python 3.10 or higher
- Pandas
- Numpy

A standard [Anaconda](https://www.anaconda.com/) installation will suffice for all of the above.

## Usage
If run as a script, will read all .xlsx files from the current working directory and use them to construct a pandas DataFrame containing call data. This DataFrame is then used to generate statistics on hourly, daily and monthly callers.
These statistics are exported to separate .xlsx files in an 'output' folder created inside the current working directory.

## future-wip
Currently working on data visualisation through matplotlib.
