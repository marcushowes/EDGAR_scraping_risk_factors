import pandas as pd


def write_dataframe_to_excel(df, file_path):

    # Write the DataFrame to Excel
    df.to_excel(file_path, index=False)

def add_row_to_output(df, row_data):

    # Create a new DataFrame with the row_data
    new_row = pd.DataFrame(row_data, index=[0])

    # Append the new row to the original dataframe
    df = pd.concat([df, new_row], ignore_index=True)

    return df

