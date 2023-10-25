import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
from openpyxl.workbook.workbook import Workbook
import streamlit as st
import requests
from streamlit_lottie import st_lottie


col3,col4 = st.columns(2)
with col3:
    st.header('Delta...')
with col4:
    ticker = st.text_input('Ticker')

if ticker:
    ticker = ticker
else:
    ticker = "Not_given"

col1,col2 = st.columns(2)
with col1:
    AR = st.file_uploader("Level 2 File, (ex- FR File)")

with col2:
    FR = st.file_uploader('Level 1 File, (ex- Analyst File)')


def load_lottiurl(url: str):
        r = requests.get(url)
        if r.status_code != 200:
            return None
        
        return r.json()
#GIF loading

loader = load_lottiurl('https://lottie.host/289ca56b-6dbb-4337-b488-895f72a1c7cb/FpIA3aCqcm.json')
done_gif = load_lottiurl('https://lottie.host/43869007-4076-48ce-8d31-c9298325d54d/4JouEu0HdT.json')
error_gif = load_lottiurl('https://lottie.host/872f9d6e-08cb-4beb-831e-2a03ae581c90/AtWtDSyzhh.json')


# For fetching all SRC from hyerlinks
def extract_hyperlinks_from_excel(excel_file):
    # Create a dictionary to store the hyperlinks.
    #src_num_dict = {}
    list_of_dict = {}

    # Load the Excel file.
    workbook = load_workbook(excel_file, data_only=True)
    
    # Iterate through all sheets in the workbook.
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Iterate through all cells in the sheet.
        for row in sheet.iter_rows():
            for cell in row:
                # Check if the cell contains a hyperlink.
                if cell.hyperlink is not None:
                    src_num_dict = {}
                    cell_location = cell.coordinate  # Cell location like 'A1'
                    src_num = cell.hyperlink.target.split("/")[-1]  # The hyperlink URL
                    src_num_dict[src_num] = [src_num, cell_location]
     
                    list_of_dict[src_num]=src_num_dict

    return list_of_dict

#SRC for units and periods
def unit_period_dict(wb_fn,deleted_src,data_added_src,AR_src):
    unit_and_period_data ={}
    for row in range (1,(wb_fn.max_row+1)):
        cell_unit = wb_fn.cell(row = row, column= 3)
        cell_period = wb_fn.cell(row = row, column= 4)
        latest_cell = wb_fn.cell(row = row, column= wb_fn.max_column)
        if cell_unit.value is not None:
            if latest_cell.value is None:
                for column in range ((-1*wb_fn.max_column),4):
                    if column < 0:
                        iter_cell = wb_fn.cell(row = row, column= (-1*column))
                    elif column > 0:
                        iter_cell = wb_fn.cell(row = row, column= (column))
                        #print(iter_cell.coordinate)
                    if iter_cell.hyperlink is not None:
                        if (iter_cell.hyperlink.target.split("/")[-1]) in deleted_src:
                            print(iter_cell.coordinate)
                            pass
                        elif (iter_cell.hyperlink.target.split("/")[-1]) in data_added_src:
                            pass
                        
                        elif (iter_cell.hyperlink.target.split("/")[-1]) in AR_src:
                            
                            unit = cell_unit.value
                            period = cell_period.value
                            src_num_dict = {}
                            cell_location = iter_cell.coordinate  # Cell location like 'A1'
                            src_num = iter_cell.hyperlink.target.split("/")[-1]  # The hyperlink URL
                            src_num_dict[src_num] = [unit,period, cell_location,iter_cell]
                            unit_and_period_data[src_num]=src_num_dict
                            break

            elif latest_cell.hyperlink is not None:
                for column in range ((-1*wb_fn.max_column),4):
                        if column < 0:
                            iter_cell = wb_fn.cell(row = row, column= (-1*column))
                        elif column > 0:
                            iter_cell = wb_fn.cell(row = row, column= (column))
                            #print(iter_cell.coordinate)
                            
                        if (iter_cell.hyperlink.target.split("/")[-1]) in deleted_src:
                            pass
                        
                        elif (iter_cell.hyperlink.target.split("/")[-1]) in data_added_src:
                            pass
                        
                        elif (iter_cell.hyperlink.target.split("/")[-1]) in AR_src:
                            unit = cell_unit.value
                            period = cell_period.value
                            src_num_dict = {}
                            cell_location = iter_cell.coordinate  # Cell location like 'A1'
                            src_num = iter_cell.hyperlink.target.split("/")[-1]  # The hyperlink URL
                            src_num_dict[src_num] = [unit,period, cell_location,iter_cell]
                            unit_and_period_data[src_num]=src_num_dict
                            break

    return unit_and_period_data

#merging functions
def merge_unmerg_dict(wb_fn):
    merge_unmerg ={}
    for row in range (1,(wb_fn.max_row+1)):
        cell_unit = wb_fn.cell(row = row, column= 3)
        latest_cell = wb_fn.cell(row = row, column= wb_fn.max_column)
        if cell_unit.value is not None:
            if latest_cell.value is None:
                for column in range ((-1*wb_fn.max_column),4):
                    if column < 0:
                        iter_cell = wb_fn.cell(row = row, column= (-1*column))
                    elif column > 0:
                        iter_cell = wb_fn.cell(row = row, column= (column))
                    if iter_cell.hyperlink is not None:
                        src_num_dict = {}
                        cell_location = iter_cell.coordinate  # Cell location like 'A1'
                        src_num = iter_cell.hyperlink.target.split("/")[-1]  # The hyperlink URL
                        src_num_dict[src_num] = [iter_cell.row, cell_location]
                        merge_unmerg[src_num]=src_num_dict
                        break

            elif latest_cell.hyperlink is not None:
                src_num_dict = {}
                cell_location = latest_cell.coordinate  # Cell location like 'A1'
                src_num = latest_cell.hyperlink.target.split("/")[-1]  # The hyperlink URL
                src_num_dict[src_num] = [latest_cell.row, cell_location]
                merge_unmerg[src_num]=src_num_dict
    return merge_unmerg

def All_SRC_in_ROW(wb_fn,row,data_added_src,deleted_src):
    columns = wb_fn.max_column
    cell_unit = wb_fn.cell(row = row, column= 3)
    src_list =[]
    for column in range (4,(columns+1)):
        iter_cell = wb_fn.cell(row = row, column= column)
        if iter_cell.hyperlink is not None:
            src_num = iter_cell.hyperlink.target.split("/")[-1]
            if src_num in data_added_src:
                pass
            elif src_num in deleted_src:
                pass
            else:
                src_list.append(src_num)
        else:
            pass
    return src_list

def Delta(AR_f,FR_f):
    #Loading excel and activating it
    AR_df = load_workbook(AR_f)
    FR_df = load_workbook(FR_f)
    AR_fn = AR_df.active
    FR_fn = FR_df.active

    if AR_fn.max_column == FR_fn.max_column:
        with col6:
            st_lottie(loader,height=250,width=250, key='loader')
    
    else:
        with col6:
            st_lottie(error_gif,height=175,width=175,key='error_gif')
        st.error(f"Both files number of Columns are not same.! , Level 1 columns- {FR_fn.max_column}, Level 2 columns- {AR_fn.max_column}")
        st.stop()

    excel_files = [AR,FR]

    # Create a new Excel workbook to consolidate the sheets.
    combined_workbook = openpyxl.Workbook()

    # Iterate through each Excel file and each sheet within each file.
    for excel_file in excel_files:
        workbook = openpyxl.load_workbook(excel_file)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            # Create a new sheet in the combined workbook with the same name.
            combined_sheet = combined_workbook.create_sheet(title=sheet_name)
            
            # Copy data from the original sheet to the combined sheet.
            for row in sheet.iter_rows():
                for cell in row:
                    combined_sheet[cell.coordinate] = cell.value

    # Remove the default sheet created by openpyxl.
    combined_workbook.remove(combined_workbook.active)

    # Save the combined workbook to a new file.
    combined_workbook.save("combined_excel.xlsx")


    combined_wb = load_workbook("combined_excel.xlsx")
    AR_sheet = combined_wb['Sheet1']
    FR_sheet = combined_wb['Sheet11']

    #created new sheet to enter count
    combined_wb.create_sheet(title="Delta")

    #changing Sheet names
    AR_sheet.title = "Level 2 file"
    FR_sheet.title = "Level 1 file"

    delta_sheet = combined_wb['Delta']

    #entering names in new sheet
    delta_sheet.cell(1,1).value='Errors'
    delta_sheet.cell(1,2).value='Count'
    delta_sheet.cell(2,1).value='Data Deleted'
    delta_sheet.cell(3,1).value='Data Added'
    delta_sheet.cell(4,1).value='Unit Error'
    delta_sheet.cell(5,1).value='Period Error'
    delta_sheet.cell(6,1).value='Merging Error'

    AR_src = extract_hyperlinks_from_excel(AR_f)
    FR_src = extract_hyperlinks_from_excel(FR_f)

    deleted_src = [item for item in FR_src if item not in AR_src]
    data_added_src = [item for item in AR_src if item not in FR_src]

    delta_sheet.cell(2,2).value= int(len(deleted_src))
    delta_sheet.cell(3,2).value=int(len(data_added_src))

    for row in range(1,FR_fn.max_row + 1):
      for column in range (1,FR_fn.max_column + 1):
        cell = FR_fn.cell(row = row, column= column)
        if cell.hyperlink is not None:
            if (cell.hyperlink.target.split("/")[-1]) in deleted_src:
                ro = cell.row
                col= cell.column
                cell_col = FR_sheet.cell(ro,col)
                cell_col.fill = PatternFill(start_color="FF0000",fill_type="solid")
                cell_col.comment = Comment("Data deleted in AR file", author="R. Praveen")

    for row in range(1,AR_fn.max_row + 1):
      for column in range (1,AR_fn.max_column + 1):
        cell = AR_fn.cell(row = row, column= column)
        if cell.hyperlink is not None:
            if (cell.hyperlink.target.split("/")[-1]) in data_added_src:
                ro = cell.row
                col= cell.column
                cell_col = AR_sheet.cell(ro,col)
                #print(cell_col.coordinate)
                cell_col.fill = PatternFill(start_color="FF0000",fill_type="solid")
                cell_col.comment = Comment("Data Added in AR file", author="R. Praveen")

    UP_dict_ar = unit_period_dict(AR_fn,deleted_src,data_added_src,AR_src)
    UP_dict_fr = unit_period_dict(FR_fn,deleted_src,data_added_src,AR_src)

    unit_count=0
    for item in UP_dict_fr:
        if item in UP_dict_ar:
            if UP_dict_fr[item][item][0] == UP_dict_ar[item][item][0]:
                pass
            else:
                ro=(UP_dict_ar[item][item][3]).row
                cell_col = AR_sheet.cell(ro,3)
                unit_count = unit_count+1
                cell_col.fill = PatternFill(start_color="FF0000",fill_type="solid")
                note = f'Unit is Changed from {UP_dict_fr[item][item][0]} to {UP_dict_ar[item][item][0]}'
                cell_col.comment = Comment(note, author="R. Praveen")
                print(f'Unit changed in FR file at {UP_dict_fr[item][item][0]}, & in AR file {UP_dict_ar[item][item][0]}')

    delta_sheet.cell(4,2).value = int(unit_count)

    period_count = 0
    for item in UP_dict_fr:
        if item in UP_dict_ar:
            if UP_dict_fr[item][item][1] == UP_dict_ar[item][item][1]:
                pass
            else:
                ro=(UP_dict_ar[item][item][3]).row
                cell_col = AR_sheet.cell(ro,4)
                period_count = period_count+1
                cell_col.fill = PatternFill(start_color="FF0000",fill_type="solid")
                note = f'Period is Changed from {UP_dict_fr[item][item][1]} to {UP_dict_ar[item][item][1]}'
                cell_col.comment = Comment(note, author="R. Praveen")
                print(f'Period changed in FR file at {UP_dict_fr[item][item][1]}, & in AR file {UP_dict_ar[item][item][1]}')

    delta_sheet.cell(5,2).value = int(period_count)

    #Merging
    MER_ar = merge_unmerg_dict(AR_fn)
    MER_fr = merge_unmerg_dict(FR_fn)

    row_vise_src_FR = {}
    row_vise_src_AR = {}
    for item in MER_fr:
        if item in MER_ar:
            row_list_fr = []
            row_list_ar = []
            row_fr = MER_fr[item][item][0]
            row_ar = MER_ar[item][item][0]
            row_list_fr = All_SRC_in_ROW(FR_fn,row_fr,data_added_src,deleted_src)
            row_list_ar = All_SRC_in_ROW(AR_fn,row_ar,data_added_src,deleted_src)
            row_vise_src_FR[item] = row_list_fr
            row_vise_src_AR[item] = row_list_ar

    #list for merging count
    Merging_count = []
    for item in row_vise_src_FR:
        if item in row_vise_src_AR:
            if row_vise_src_FR[item] == row_vise_src_AR[item]:
                pass
            else:
                row = MER_fr[item][item][0]
                row_ar = MER_ar[item][item][0]
                ar_count=len(row_vise_src_AR[item])
                fr_count = len(row_vise_src_FR[item])
                if ar_count > fr_count:
                    final_count = ar_count
                else:
                    final_count = fr_count
                print(f'in FR file row no- {row}, was changed in AR file. Row in AR file {row_ar}. count - {final_count}')
                Merging_count.append(final_count)
                row = AR_sheet[row_ar]
                for cell in row:
                    cell.fill = PatternFill(start_color="FF0000",fill_type="solid")
                    note = f'Merging Error was corrected in this row'
                    cell.comment = Comment(note, author="R. Praveen")

    delta_sheet.cell(6,2).value=int(sum(Merging_count))

    combined_wb.save("combined_excel.xlsx")
        

d_but = st.button("Delta Review")
col5,col6 = st.columns(2)
download = False
if d_but:
    with st.spinner("Reviewing...."):
        Delta(AR,FR)   
    download = True

data = 'combined_excel.xlsx'

# Read the file content
with open(data, 'rb') as file:
    file_content = file.read()

file_n = f'{ticker}_delta.xlsx'

if download:
    st.download_button("Download file",data=file_content,file_name=file_n,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
