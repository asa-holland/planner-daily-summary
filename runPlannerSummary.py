import pandas as pd


def summarize_planner_export(sample_filepath):
    print(pre_process(sample_filepath))
    return


def pre_process(infile):

    required_columns = []
    pre_processed_df = pd.read_excel(infile)
    return pre_processed_df


if __name__ == "__main__":
    SAMPLE_FILEPATH = 'Project Workflow.xlsx'
    summarize_planner_export(SAMPLE_FILEPATH)
