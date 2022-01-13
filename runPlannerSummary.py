import pandas as pd
from datetime import date, datetime, timedelta
pd.set_option('display.max_columns', None)


def summarize_planner_export(sample_filepath):
    pre_process_result = pre_process(sample_filepath)
    post_process_result = post_processing(pre_process_result)
    format_final_result(post_process_result)
    return

def printCols(df):
    print(df.columns)
def pre_process(infile):

    # default columns
    # ['Task ID', 'Task Name', 'Bucket Name', 'Progress', 'Priority',
    #  'Assigned To', 'Created By', 'Created Date', 'Start Date', 'Due Date',
    #  'Late', 'Completed Date', 'Completed By', 'Description',
    #  'Completed Checklist Items', 'Checklist Items', 'Labels']
    # ['Task ID', 'Task Name', 'Bucket Name', 'Progress', 'Priority',
    #  'Assigned To', 'Created By', 'Created Date', 'Start Date', 'Due Date',
    #  'Late', 'Completed Date', 'Completed By', 'Description',
    #  'Completed Checklist Items', 'Checklist Items', 'Labels']

    required_columns = ['Task Name', 'Priority', 'Assigned To', 'Due Date', 'Late', 'Description']
    df = pd.read_excel(infile, usecols=required_columns)

    print(df['Due Date'].head())
    # "01/10/2022"

    # convert due date from sting to datetime
    # df['Due Date'] = pd.to_datetime(df['Due Date']).apply(lambda x: datetime.strftime(x, '%y-%m-%d'))
    # df['Due Date'] = df['Due Date'].dt.date
    # df['Due Date'] = df['Due Date'].apply(lambda x: datetime.date(x.year, x.month, x.day))

    # set up string of today and tomorrow date
    today = str(date.today().strftime("%m/%d/%Y"))
    tomorrow = date.today() + timedelta(days=1)
    tomorrow = tomorrow.strftime("%m/%d/%Y")

    print(type(today), today)
    print(type(tomorrow), tomorrow)

    print(df["Due Date"].unique())

    due_today = df["Due Date"] == today
    due_tomorrow = df["Due Date"] == tomorrow

    # print(type(df["Due Date"][0]), df["Due Date"][0])
    print(due_today.unique())
    pre_processed_df = df.loc[due_today]

    print('HERE', df.loc[due_tomorrow])

    # printCols(pre_processed_df)
    return df


def post_processing(dataframe):

    # Add Categories column
    post_processed_dataframe = dataframe

    # Remove Due Date Columns
    post_processed_dataframe.drop('Due Date', 1)

    # TODO: remove populate Category column and drop Description column

    # Create custom sort order
    df_urgency_order = pd.DataFrame({
        'urgency': ['Urgent', 'Important', 'Medium', 'Low'],
    })
    sort_urgency = df_urgency_order.reset_index().set_index('urgency')

    # Create new column for sort order
    post_processed_dataframe['urgency_order'] = post_processed_dataframe['Priority'].map(sort_urgency['index'])

    # Sort by urgency_order
    post_processed_dataframe = post_processed_dataframe.sort_values('urgency_order')

    # then by Priority using custom sort 'Urgent', 'Important', 'Medium', 'Low'
    # post_processed_dataframe = post_processed_dataframe.sort_values('Priority')

    return post_processed_dataframe


def format_final_result(dataframe):
    print(dataframe)

    # export to excel file

    # bold top row

    return

if __name__ == "__main__":
    SAMPLE_FILEPATH = 'Project Workflow.xlsx'
    summarize_planner_export(SAMPLE_FILEPATH)
