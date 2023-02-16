import os
import iterating_two_files
import making_excel
import pandas as pd


# Paste folderpath here
folderpath = r'C:\Users\Marcus.Howes_PLA\Desktop\testing'
cik_excel_path = r'C:\Users\Marcus.Howes_PLA\Desktop\ciks.xlsx'
df_ciks = pd.read_excel(cik_excel_path, usecols=[0])
fail_count = 0

df_fail = pd.DataFrame(columns=['Number', 'File Name', 'Reason'])
fail_count = 1

df_out = pd.DataFrame(columns=['CIK', 'Years', 'Similarity', 'Longer', 'Difference', 'Later Publish Date', 'Length 1', 'Length 2'])
max_file_size = 250 # maximum file size in KB


for root, dirs, files in os.walk(folderpath):

    for dirname in dirs:

        print('At subdir_path')
        # Do something with each subdirectory
        subdir_path = os.path.join(root, dirname)

        last_number = subdir_path.split("\\")[-1].strip()
        is_present = df_ciks.isin([int(last_number)]).values.any()

        if is_present == False:
            # Add the file to the failed files excel.
            reason = 'CIK not found'
            df_fail = making_excel.add_row_to_output(df_fail, {'Number': fail_count, 'File Name': last_number, 'Reason': reason})
            fail_count = fail_count + 1   
            continue  

        #Skip if the folder is an excel file
        if any(file.endswith('.xlsx') for file in os.listdir(subdir_path)):
            continue

        for filename in os.listdir(subdir_path):
            file_path = os.path.join(subdir_path, filename)

            if os.path.isfile(file_path) and filename.startswith('risk_factors'):
                size_kb = os.path.getsize(file_path) / 1024

                if size_kb > max_file_size:

                    reason = 'File too big'
                    df_fail = making_excel.add_row_to_output(df_fail, {'Number': fail_count, 'File Name': filename, 'Reason': reason})
                    fail_count = fail_count + 1 

                    os.remove(file_path)


        # Call the function to process the files that come a year after the other
        df_out = iterating_two_files.process_files(subdir_path, df_out) 

excel_out_file_path = f'{folderpath}\output.xlsx'
making_excel.write_dataframe_to_excel(df_out, excel_out_file_path)

excel_fail_file_path = f'{folderpath}\output_failed_files.xlsx'
making_excel.write_dataframe_to_excel(df_fail, excel_fail_file_path)

print(f"\n DONE! \n")
