import pandas as pd
from openpyxl.styles import PatternFill
from datetime import datetime

# Get the current date and time
now = datetime.now()

# Format the date and time
date_time = now.strftime("%m-%d-%Y")

from openpyxl.styles import Font
#file1 : old file - teams site
#file2 : new file - ddf

file1 = pd.ExcelFile(r"C:\Users\eg724520\Valmont Industries, Inc\Global IT AppSOD & DP - Documents\Team Tools\PyhtonCode - SAP Workstream DDF Data Refresh\Process Step Descriptions\Output\sync_results_ProcessStepDesc_09-19-2023.xlsx")

file2 = pd.ExcelFile(r"C:\Users\eg724520\Valmont Industries, Inc\Global IT AppSOD & DP - Documents\Team Tools\PyhtonCode - SAP Workstream DDF Data Refresh\Process Step Descriptions\Input\ProcessStepDesc_1680093583905_Excel.xlsx (7).xlsx")
sheet_names = file1.sheet_names

#sync_path is the output path, modified it if it's needed
sync_path = fr"C:\Users\eg724520\Valmont Industries, Inc\Global IT AppSOD & DP - Documents\Team Tools\PyhtonCode - SAP Workstream DDF Data Refresh\Process Step Descriptions\Output\sync_results_ProcessStepDesc_{date_time}.xlsx"
# sync_path = r"C:\Users\eg724520\OneDrive - Valmont Industries, Inc\Documents\DDFCompare\sync_results_level3trackingdata.xlsx"

#key -> column in the file that can act as the primary key of the data.
#this key can be added if dealing with a new excel file in the future
key = {'L3_Process_ID', 'Busines Process L3', 'Level 3 Business Process Process Id','T_Code', 'Name'}


highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
with pd.ExcelWriter(sync_path, engine='openpyxl') as writer:
    for sheet_name in sheet_names:
        key_columns = []
        try:
            sheet1 = file1.parse(sheet_name)
            sheet2 = file2.parse(sheet_name)

            comment_dict = {}

            for col in key:
                if col in sheet2.columns:
                    key_columns.append(col)

            print(f"Processing sheet: {sheet_name}, Key Columns:", end = '')
            for common_col in key_columns:
                print(f" {common_col},", end = '')

            for index, right_row in sheet2.iterrows():
                try:
                    for x, left_row in sheet1.iterrows():
                        count = 0
                        if all(left_row[key] == right_row[key] for key in key_columns):

                            for comment in sheet1.columns:

                                if comment not in sheet2.columns:
                                    comment_dict[count] = comment

                                count += 1
                            for key_index, values in comment_dict.items():

                                if right_row[key_columns[0]] == left_row[key_columns[0]] and values not in sheet2.columns:
                                    sheet2.insert(key_index, values, left_row[values]) #based on the key from the old and new file, the comment or added column in the old file will be brought to the new file
                                elif right_row[key_columns[0]] == left_row[key_columns[0]] and values in sheet2.columns:
                                    sheet2.at[index, values] = left_row[values]

                except Exception as e:
                    print(f"Error processing row {index} in sheet '{sheet_name}': {e} ")

            sheet2.to_excel(writer, sheet_name=f'{sheet_name}', index=False)
            styled_sheet = writer.sheets[f'{sheet_name}']
            print(f"Join {sheet_name} successfully")
            for cell in styled_sheet[1]:
                if cell.value in comment_dict.values():
                    # Apply the desired font formatting to the cell
                    cell.fill = highlight_fill
        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")



