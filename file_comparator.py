''' Module to compare two excel-type files and to highlight the differences '''

import argparse
import sys
import os
import pandas as pd

def row_iterator(df, cols):
    '''  Read the given DataFrame row-wise and returns a list of tuples 
         with each tuple containing a single record.
    '''
    rows = []
    for _,row in df.iterrows():
        r = tuple(row[cols])
        rows.append(r)
    return rows

def rows_highlighter(x,colum):
    ''' The new/updated/deleted records are being highlighted based on the condition. 
    '''
    if x.colors_to_identify == 'yellow':
        return ['background-color: yellow']*len(colum)
    elif x.colors_to_identify == 'green':
        return ['background-color: green']*len(colum)
    elif x.colors_to_identify == 'red':
        return ['background-color: red']*len(colum)
    else:
        return ['background-color: white']*len(colum)    

def columns_highlighter(s):
    ''' Highlight entire columns'''
    return 'background-color: % s' % 'green'
   
def values_evaluator(base_df, tar_df, target):
    ''' Compare each record row-wise and identify the newly added/ updated / deleted 
        rows and columns.
    '''
    base_col = base_df.columns 
    tar_col = tar_df.columns
    
    ## Comparing the columns ##
    missed_cols = [col for col in base_col if col not in tar_col]
    common_cols = [col for col in base_col if col in tar_col]
    extra_cols = [col for col in tar_col if col not in base_col]
    
    tar_pf = tar_df[common_cols]
    tar_pf = tar_pf.fillna('')
    base_pf = base_df[common_cols]
    base_pf = base_pf.fillna('')
    
    ## Extract row values ##
    base_rows = row_iterator(base_pf, common_cols)
    tar_rows = row_iterator(tar_pf, common_cols)
    
    ## Identify common, missed and extra rows ##
    miss_rows = []
    comm_rows = []
    for row in base_rows:
        if row in tar_rows:
            comm_rows.append(row)
        else:
            miss_rows.append(row)
    extra_rows = [row for row in tar_rows if row not in base_rows]
    
    ## Logic to find Deleted rows and Updated rows##
    new_rows = []
    upd_rows = []
    indexes = []
    tr = [mr[0] for mr in miss_rows]
    for row in extra_rows: 
        if row[0] in tr:
            tr_count = [pos for pos, rr in enumerate(tr) if rr == row[0]]
            idx = tr_count[0]
            occur = 1
            for trc in tr_count:
                mr_set = set(miss_rows[trc])
                row_set = set(row)
                val = mr_set & row_set
                if len(val) > occur:
                    occur = len(val)
                    idx = trc
            if idx not in indexes:
                indexes.append(idx)
                upd_rows.append(row) 
            else:
                upd_rows_len = len(upd_rows)
                for trc in tr_count:
                    if trc not in indexes:
                        indexes.append(idx)
                        upd_rows.append(row)
                        break
                if len(upd_rows) != upd_rows_len+1:
                    new_rows.append(row)  
        else:
            new_rows.append(row)
    del_rows = [row for idx,row in enumerate(miss_rows) if idx not in indexes]
    
    ### generate output ###
    out_rows = {}
    output_sam = []
    colors = []
    for idx, row in enumerate(tar_rows):
        val = {"row" : list(row)}
        if row in upd_rows:
            val.update({'colors_to_identify': 'yellow'})
        elif row in new_rows:
            val.update({'colors_to_identify': 'green'})
        else:
            val.update({'colors_to_identify': 'white'})
        colors.append(val['colors_to_identify'])
        output_sam.append(val['row'] + [val['colors_to_identify']])
        out_rows.update({str(idx):val})
    col = common_cols + ['colors_to_identify']
    
    output_df = pd.DataFrame(output_sam, columns = col)
    tar_df['colors_to_identify'] = colors
    
    ### highlighting rows ###
    tar_row_df = tar_df.style.apply(rows_highlighter, colum= col, subset = col, axis=1)
    
    ### highlighting columns ###
    tar_col_df = tar_row_df.applymap(columns_highlighter, subset = pd.IndexSlice[:, extra_cols])
    
    summary_dict = [
        {
            "Type"  : "Total Records",
            "count" : len(tar_df),
            "colors_to_identify": "white"
        },
        {
            "Type" : "Updated Records",
            "count": len([c for c in colors if c == 'yellow']),
            "colors_to_identify": "yellow"
        },
        {
            "Type" : "New Records",
            "count": len([c for c in colors if c == 'green']),
            "colors_to_identify": "green"
        },
        {
            "Type" : "Deleted Records",
            "count": len(del_rows),
            "colors_to_identify": "red"
        },
        {
            "Type" : "New Columns",
            "count": len(extra_cols),
            "colors_to_identify": "green"
        },
        {
            "Type" : "Deleted Columns",
            "count": len(missed_cols),
            "colors_to_identify": "red"
        }
    ]
    summ_df = pd.DataFrame(summary_dict)
    sum_col = list(summ_df.columns)
    summary_df = summ_df.style.apply(rows_highlighter, colum=sum_col , subset = sum_col, axis=1)
    
    ## Writing the final output ##
    out_filename = 'output_'+target
    writer = pd.ExcelWriter(out_filename, engine='xlsxwriter')
    final_col = list(tar_df.columns)
    final_col.remove('colors_to_identify')
    tar_col_df.to_excel(writer, sheet_name='output', columns=final_col, index=False)
    
    summary_df.to_excel(writer, sheet_name='summary', columns=['Type','count'], index=False)
    writer.save()
    writer.close()
    print(f"Output file generated : {out_filename}")
    
def file_extension_checker(file):
    ''' Check whether the given file is of extension .xlsx/.xls/.csv
    '''
    if file.endswith('.xlsx') or file.endswith('.xls'):
        df = pd.read_excel(file)
        df = df.drop_duplicates() # Duplicates Removed
    elif file.endswith('.csv'):
        df = pd.read_csv(file)
        df = df.drop_duplicates() # Duplicates Removed
    else:
        df= pd.DataFrame()
        
    return df

def tool_display():
    ''' To Display the Tool logo
    '''
    print("         -----------------------------------------------------------------------------------------")
    print("         ----------  ___          ___      ___  __          ___    _    ___   ___   ___  ---------")
    print("         ---------- |    |  |    |        |    |  | |\  /| |   |  / \  |   | |     |   | ---------")
    print("         ---------- |___ |  |    |___     |    |  | | \/ | |___| /___\ |___| |___  |___| ---------")
    print("         ---------- |    |  |    |        |    |  | |    | |     |   | | \   |     | \   ---------")
    print("         ---------- |    |  |___ |___     |___ |__| |    | |     |   | |  \  |___  |  \  ---------")
    print("         ----------                                                                      ---------")
    print("         -----------------------------------------------------------------------------------------")
    print("         Version : 0.1 - 1st Feb 2022")
    print("         Created By: Promoth  https://github.com/linacpromoth")
    print("         -----------------------------------------------------------------------------------------")

def files_checker(base, target):
    ''' Check whether base file and target file are present
    '''
    if os.path.exists(base) and os.path.exists(target):
        base_df = file_extension_checker(base)
        tar_df = file_extension_checker(target)
        if len(base_df) != 0 and len(tar_df) != 0:
            tool_display()
            
            ### comparison ###
            values_evaluator(base_df, tar_df, target)
        elif  len(tar_df) == 0:  
            print(f"Target file :{target} is empty")
            sys.exit()
        elif len(base_df) == 0:  
            print(f"Base file :{base} is empty")
            sys.exit()
        else:
            print("Ensure both base and target file contains any one of the following extensions")
            print("1. .xlsx\n2. .xls\n3. .csv")
            sys.exit()
    else:
        print("Ensure Both The Files Are Present.\nFiles Not Found!!!!")
        sys.exit()
    
    

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Compare target file with base file')
    parser.add_argument("-b", "--base_file", help="the base file against which comparison takes place", required=True)
    parser.add_argument("-t", "--target_file", help="the target file to be compared with", required=True)
    
    if len(sys.argv)!=5:
        parser.print_help(sys.stderr)
        sys.exit()
    
    args = parser.parse_args()
    BASE = args.base_file
    TARGET = args.target_file
    
    ## Check presence of files ##
    files_checker(BASE, TARGET)
    
    


    