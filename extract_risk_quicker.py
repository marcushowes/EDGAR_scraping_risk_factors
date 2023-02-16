import os
from bs4 import BeautifulSoup
import re
import lxml.html
import io
import html2text
import iterating_two_files
import making_excel
import pandas as pd


def read_file_in_chunks(filepath, chunk_size=1*1024*1024, max_size=10*1024*1024):
    with open(filepath, "r", encoding='utf-8') as f:
        raw_html = ""
        chunk = f.read(chunk_size)
        while chunk and len(raw_html) + len(chunk) <= max_size:
            raw_html += chunk
            chunk = f.read(chunk_size)
    return raw_html

part_pattern = re.compile("(?s)(?i)(?m)> +Part|>Part|^Part", re.IGNORECASE + re.MULTILINE)
item_pattern = re.compile("(?s)(?i)(?m)> +Item|>Item|^Item", re.IGNORECASE + re.MULTILINE)
substitute_html = re.compile("(?s)<.*?>")

# Paste folderpath here
folderpath = r'C:\Users\Marcus.Howes_PLA\Desktop\testing'
cik_excel_path = r'C:\Users\Marcus.Howes_PLA\Desktop\ciks.xlsx'
df_ciks = pd.read_excel(cik_excel_path, usecols=[0])
fail_count = 0


# ---------------- FILE ITERATION AND RISK FACTOR SCRAPE ----------------

for root, dirs, files in os.walk(folderpath):
    for filename in files:
        filepath = os.path.join(root, filename)

        if os.path.isfile(filepath):

            raw_html = read_file_in_chunks(filepath)

            # search for the parts/items of filing and replace it with unique character to split on later
            updated_html = part_pattern.sub(">°Part", raw_html)
            updated_html = item_pattern.sub(">°Item", updated_html)

            # remove tables because they can be parsed seperately
            lxml_html = lxml.html.fromstring(updated_html)
            base = lxml_html.getroottree()

            # remove tables because we can analyze them seperately
            for i, table in enumerate(base.iter(tag='table')):

                table_text = table.text_content()

                # i just used two entries to determine whether we were looking at toc table
                if "Financial Data" in table_text or "Mine Safety Disclosures" in table_text:
                    pass
                else:
                    # drop table from HTML
                    table.drop_tree()

            updated_raw_html = lxml.html.tostring(base)
            soup = BeautifulSoup(updated_raw_html, 'lxml')
            h = html2text.HTML2Text()
            raw_text = h.handle(soup.prettify())

            combined_text = ""
            file_count = 0

            for idx, item in enumerate(raw_text.split("°Item")):
                if "risk factors" in item.lower():
                    if len(item) > 100:
                        first_line = item.splitlines()[0].strip()
                        
                        if "risk factors" in first_line.lower():
                            file_count += 1
                            combined_text += item

            if file_count > 0:
                filename = f"risk_factors_{filename}"
                with io.open(os.path.join(root, filename), "w", encoding='utf-8') as f:
                    f.write(combined_text)
                
                print(f"Risk Factors Sections Saved:\n{folderpath}, {filename}")

            else:
                # Making a df with all the failed files
                if fail_count < 1:
                    df_fail = pd.DataFrame(columns=['Number', 'File Name', 'Reason'])
                    fail_count = 1
                    reason = 'Risk Factors not extracted'
                    df_fail = making_excel.add_row_to_output(df_fail, {'Number': fail_count, 'File Name': filename, 'Reason': reason})

                else:
                    df_fail = making_excel.add_row_to_output(df_fail, {'Number': fail_count, 'File Name': filename, 'Reason': reason})
                    fail_count = fail_count + 1

if fail_count < 1:
    df_fail = pd.DataFrame(columns=['Number', 'File Name', 'Reason'])
    fail_count = 1

# ---------------- COMPARE SIMILARITY INTO EXCEL ----------------

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
