def import_call_data(path = None, input_hours = ('09:00', '20:59')):
    '''Returns a pandas DataFrame containing call data.
    
    Takes a path as argument. All .xlsx files in the supplied directory are imported into a dataframe.
    The user SHOULD therefore make sure the directory only contains relevant files.
    Default path is set to os.getcwd().'''

    # We manually set the path argument to os.getcwd() if none is supplied.
    if path == None:
        from os import getcwd
        path = getcwd()


    def raw_import(p = path):
        '''Returns a pandas DataFrame containing the raw .xlsx data.
        
        Takes a path as argument. As stated in the main function, this imports all .xlsx files in said directory.
        No processing is done. The main purpose of this function is to identify all .xlsx files in the provided path and prepare a DataFrame for processing.'''

        # Get all .xlsx in current dir        
        import glob
        xlsx_files = glob.glob('*.xlsx')

        # Use list comprehension inside pd.concat() to generate df
        # We need to silence warnings because you will get one for each import. These state the workbook contains no default style. This has no negative consequences for analysis.
        import warnings
        with warnings.catch_warnings(record=True):
            warnings.simplefilter('always')
            df = pd.concat(pd.read_excel(p + '\\' + xlsx_file) for xlsx_file in xlsx_files)

        # return
        return df


    def process_raw(df, working_hours=('00:00', '23:59')):
        '''Returns a pandas DataFrame containing processed call data.
        
        Takes a pandas DataFrame as argument. This is expected to be provided by raw_import().
        The working_hours argument is used to chuck out irrelevant rows. Default is to include everything.
        This is a lot of manual editing. Refer to a raw xlsx export to cross-check why rows are dropped, what column names are picked, etc.'''


        ''' Drop first two and last three rows (from EACH of the four reports). Identified by content because index and row# for the final rows are variable.
        We use the fourth column (Unnamed: 3) because this contains the department name if it's an actual datarow. 
        Given that the department name is variable we control for empty rows and 'CSQ Name' (which is the header) instead.'''
        row_filter_1 = df.loc[:,'Unnamed: 3'].notna()
        row_filter_2 = df.loc[:,'Unnamed: 3'] != 'CSQ Name'
        df = df.loc[row_filter_1 & row_filter_2]
        
        # And drop irrelevant columns: [0, 2, 3, 13]. (Column 2 is irrelevant because it contains interval end time, which is redundant as its always +30 mins compared to start time.)
        df = df.drop(columns = df.columns[[0, 2, 3, 13]])
        
        # Set correct column names
        df.columns = [
            'Interval Start Time',
            'Total Calls Presented',
            'Presented Average Queue Time',
            'Presented Max Queue Time',
            'Total Calls Handled',
            'Average Handle Time',
            'Max Handle Time',
            'Total Calls Abandoned',
            'Abandoned Average Queue Time',
            'Abandoned Max Queue Time'
        ]

        # Fix interval times to datetime format. Thanks Cisco
        df['Interval Start Time'] = pd.to_datetime(df['Interval Start Time'], format = '%m/%d/%y %I:%M:%S %p') 

        # Set interval times to index
        df = df.set_index('Interval Start Time')

        # Set 'Total Calls X' to int
        df = df.astype({'Total Calls Presented': 'int',
                        'Total Calls Handled': 'int',
                        'Total Calls Abandoned': 'int'})
        
        # Fix queue, handle time to datetime format
        time_columns = ['Presented Average Queue Time', 
                        'Presented Max Queue Time', 
                        'Average Handle Time', 
                        'Max Handle Time', 
                        'Abandoned Average Queue Time', 
                        'Abandoned Max Queue Time']
        
        for col in time_columns:
            df[col] = pd.to_datetime(df[col], format = '%H:%M:%S').dt.time

        # Remove rows based on working hours argument. 
        df = df.between_time(working_hours[0], working_hours[1]) 

        # Remove weekends
        df = df.loc[df.index.weekday.isin(range(0,5))]

        return df
    

    # resolve
    import pandas as pd
    raw_df = raw_import()
    processed_df = process_raw(raw_df, working_hours = input_hours)
    return processed_df

def generate_call_data(df, data_type = 'd'):
    '''Returns a df with relevant call statistics.
    
    'type' argument supports the following: m(onthly), w(eekly), d(aily), w(eek)d(ay), h(ourly). 
    This provides number of handled calls per month, per week, per day, per weekday (Mon - Fri) or per half hour.'''
    


    # Define range of acceptable arguments
    type_range = { 
        'm': 'monthly',
        'w': 'weekly',
        'd': 'daily',
        'wd': 'weekday',
        'h': 'hourly'
    }
    
    # Check if actual argument was within accepted range
    if data_type not in type_range and data_type not in type_range.values():
        raise KeyError # Incorrect data_type argument 
    
    # Correct abbreviation to full name
    if len(data_type) <= 2:
        data_type = type_range[data_type]

    # Actual data generation
    def groupby_type(df, groupindex):
        import pandas as pd
        return pd.DataFrame(df.groupby(groupindex)['Total Calls Handled'].apply(sum))
    
    match data_type: # Python 3.10 and up!
        # Generate df with month as index and sum of callers over month as sole column.
        case 'monthly':
            result = groupby_type(df, df.index.month)
        
        # Generate df with week as index and sum of callers per week as sole column.
        case 'weekly':
            pass #NYI

        # Generate df with date as index and sum of callers per day as sole column.
        case 'daily':
            result = groupby_type(df, df.index.date)
            result['Rolling Average Callers Handled'] = result.rolling(4).mean()                                     

        # Generate df with weekdays as index and sum of callers over weekday as sole column.
        case 'weekday':
            result = groupby_type(df, df.index.day_name())

        # Generate df with half hours as index and sum of callers over time as sole column.
        case 'hourly':
            result = groupby_type(df, df.index.time)

    # Now we export to xlsx
    # xlsx export name structure: [timespan]_[datatype]_results.xlsx
    timespan = str(df.index.month[0]).zfill(2) + '_' + str(df.index.year[0]) + '_to_' + str(df.index.month[-1]).zfill(2) + '_' + str(df.index.year[-1])
    export_name = 'output\\' + timespan + '_' + data_type + '_' + 'results.xlsx'      

    # Check if an output directory already exists; if not, make one.
    from os import mkdir
    try:
        mkdir('output')
    except FileExistsError:
        pass

    # Export to xlsx and return
    result.to_excel(export_name)
    return result


def generate_call_graph(df, file_name):
    '''Exports a .png file with a graph describing caller data.
    
    Expects a pandas DataFrame as argument generated by cisco_statistics.generate_call_data() found in this library.'''
    import matplotlib.pyplot as plt
    plt.plot(df.index, df.iloc[:,-1]) #Take the last column because daily also has a column with nominal values but you want the rolling average. For all other inputs the last column is the only one in the df.
    #plt.plot.show()
    
    # Export to png
    plt.savefig('output\\' + file_name)


def main():
    primary_df = import_call_data()
    hourly_data = generate_call_data(primary_df, 'h')
    monthly_data = generate_call_data(primary_df, 'm')
    daily_data = generate_call_data(primary_df, 'd')
    generate_call_graph(daily_data, 'daily_data.png')

if __name__ == '__main__':
    main()