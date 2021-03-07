# Tool to compare table documents

# import modules
import pandas as pd
import numpy as np
# for isnan
import math
# for opening an excel workbook to add sheets
from openpyxl import load_workbook
# for listing available files
import os
# for writing several excel subsheets
import xlsxwriter
# time stamp
import time
# for text similarity
#import nltk
#from nltk.tokenize import word_tokenize, sent_tokenize
#from nltk.corpus import stopwords


# input of file data like file name, index...
def fcn_input_filedata():
    print("Which revisions of tables do you want to compare?")



    # ENTER working directory
    print("\n1.  Enter working directory:             ('Enter' for default: 'C:\CompareDir\')", end='')
    working_directory = input()
    if working_directory == "":
        print("    -> default set")
        working_directory = "C:\\CompareDir\\"

    print("\n    Available files:")
    arr = os.listdir(working_directory)
    print(arr)



    # ENTER file old
    print("\n2a. Enter file name of older version:    ('Enter' for default: 'old.xlsx')", end='')
    file_old = input()
    if file_old == "":
        print("    -> default set")
        file_old = "old.xlsx"
    
    # create filepath
    print("\n    Path 'Old file': ")
    file_path_old = working_directory + file_old
    print("    " + file_path_old)

    # ENTER index new
    print("\n2b. Enter index name:                    ('Enter' for default: 'Object ID')", end='')
    index_name_old = input()
    if index_name_old == "":
        print("    -> default set")
        index_name_old = "Object ID"



    # ENTER file old
    print("\n3a. Enter file name of newer version:    ('Enter' for default: 'new.xlsx')", end='')
    file_new = input()
    if file_new == "":
        print("    -> default set")
        file_new = "new.xlsx"

    # create filepath
    print("\n    Path 'New file': ")
    file_path_new = working_directory + file_new
    print("    " + file_path_new)

    # ENTER index new
    print("\n3b. Enter index name:                    ('Enter' for default: 'Object ID')", end='')
    index_name_new = input()
    if index_name_new == "":
        print("    -> default set")
        index_name_new = "Object ID"

    return(file_path_old, index_name_old, file_path_new, index_name_new, working_directory)



def fcn_read_comparefiles(file_path_old, index_name_old, file_path_new, index_name_new):
    # open tables and read to data frames
    try:
        df_old = pd.read_excel(file_path_old,"Tabelle1",index_col=index_name_old).fillna(0)
        print("\nRevision \"old\" read successfully.")
        #df_old.to_excel("Output_Old.xlsx",'Old')    
    except:
        print("An error occured, reading comparison file A.")

    try:
        df_new = pd.read_excel(file_path_new,"Tabelle1",index_col=index_name_new).fillna(0)
        print("Revision \"new\" read successfully.")
        #df_new.to_excel("Output_New.xlsx",'New')
    except:
        print("An error occured, reading comparison file B.")

    return (df_old, df_new)



def fcn_columns_createlists(df_old, df_new):
    # attributes of old and new data frame
    cols_old = df_old.columns
    cols_new = df_new.columns

    # attributes, common between old and new -> comparable
    cols_common = list(set(cols_old).intersection(cols_new))
    cols_deleted = list(set(cols_old).difference(cols_new))
    cols_added = list(set(cols_new).difference(cols_old))

    print("\n>>> Common attributes:")
    print(cols_common)
    print("\n>>> Deleted attributes:")
    print(cols_deleted)
    print("\n>>> Added attributes:")
    print(cols_added)

    return(cols_old, cols_new, cols_common, cols_added, cols_deleted)


def fcn_simiText(value_old,value_new):

    simiText = "to be implemented"

    return(simiText)


# main program
def main():
    # welcoming
    print("\n")
    print("-------------------------------")
    print(">>>   Welcome to CompareX   <<<")
    print("-------------------------------")
    print("Feedback to arndt.seb@gmail.com")
    print("\n")


    # call fcn to read file names, indices...
    file_path_old, index_name_old, file_path_new, index_name_new, working_directory = fcn_input_filedata()


    # call fcn to read data frames from comparison files
    df_old, df_new = fcn_read_comparefiles(file_path_old, index_name_old, file_path_new, index_name_new)


    # create different lists of columns for further processing
    cols_old, cols_new, cols_common, cols_added, cols_deleted = fcn_columns_createlists(df_old, df_new)
    

    
    print("\n\nAttribute comparison")
    # all values as objects to avoid errors while cycling through columns
    for conv_cols in cols_new:
        df_new[conv_cols] = df_new[conv_cols].astype(object)
    for conv_cols in cols_old:
        df_old[conv_cols] = df_old[conv_cols].astype(object)


    # data frame with diffences, based on new data frame
    df_new_diff = df_new.copy()
    df_new_diff.insert(0,"HardDiff",np.nan)
    df_new_diff.insert(1,"SoftDiff",np.nan)
    df_new_diff.insert(2,"Changes",np.nan)

    # run through new data frame to check for 
    #  - differences (case SAME / DIFF)
    #  - added rows (case ADD)
    #  (case decision whithin loop; deleted rows in separate loop)

    stats_same = 0
    stats_changed = 0
    stats_added = 0
    stats_removed = 0

    index_new = df_new.index

    for c_row in index_new:
        # - If sample is in old dataframe (in addtion zu new datafram), then it is comparable
        if (c_row in df_old.index): #and (c_row in df_new.index):
            # IMPORTANT: Set "SAME" here already, before the columns are being run through!
            #            In the column loop, only one "diff" is enough, to set the complete sample to diff.
            df_new_diff.loc[c_row,"HardDiff"] = ("SAME")

            # detailed cell check by running through common attributes (columns)
            for c_col in cols_common:
                #print(c_row)
                value_old = df_old.loc[c_row,c_col]
                value_new = df_new.loc[c_row,c_col]
                # Initialize ComparisonResult entry with "same". "Diff" will be set if one different cell is detected. 
                
                # CALL soft difference function
                if c_col == "Beschreibung":
                    simiText = fcn_simiText(value_old,value_new)
                    df_new_diff.loc[c_row,"SoftDiff"] = simiText

                # CHECK hard difference
                # - SAME
                if (value_old == value_new) or (value_new=="0" and value_old=="0"): #or (value_old is np.nan and value_new is np.nan):
                    df_new_diff.loc[c_row,c_col] = value_new
                    stats_same = stats_same + 1

                # - DIFF
                else:
                    df_new_diff.loc[c_row,c_col] = ("{} → {}").format(value_old,value_new)
                    
                    if pd.isnull(df_new_diff.loc[c_row,"Changes"]):
                        df_new_diff.loc[c_row,"Changes"] = c_col                    
                    else:
                        df_new_diff.loc[c_row,"Changes"] = df_new_diff.loc[c_row,"Changes"] + ", " + c_col

                    df_new_diff.loc[c_row,"HardDiff"] = ("DIFF")
                    stats_changed = stats_changed + 1
          
        # - ADD
        # if sample is not in old dataframe (but in new one), it was newly added
        else:
            df_new_diff.loc[c_row,"HardDiff"] = ("ADD")
            stats_added = stats_added + 1


    # run through old data frame to check for
    #  - deleted rows (DEL)
    for c_row_old in df_old.index:

        # - DEL
        if c_row_old not in df_new.index:
            df_new_diff = df_new_diff.append(df_old.loc[c_row_old,:])
            df_new_diff.loc[c_row_old,"HardDiff"] = ("DEL")
            stats_removed = stats_removed + 1


    # Add markers to header for (A)dded or (D)eleted column
    # run through header
    for c_col in df_new_diff.columns:
        if c_col in cols_added:
            df_new_diff = df_new_diff.rename(columns={c_col: 'ADD_' + c_col})
        elif c_col in cols_deleted:
            df_new_diff = df_new_diff.rename(columns={c_col: 'DEL_' + c_col})


    print("\n\nOutputs")

    print("\n\n>>> Old revision (head only)")
    print(df_old.head(3))

    print("\n\n>>> New revision (head only)")
    print(df_new.head(3))

    print("\n\n>>> New revision with differences (head only)")
    print(df_new_diff.head(3))
    
    #df_new_diff.to_excel("Output_Diff.xlsx",'Diff')
    #df_new_diff.to_excel("Output_Diff.xlsx",'Diff1')

    print(stats_same)

    df_stats = pd.DataFrame(columns=['Same','Changed','Added','Removed'])
    df_stats.loc[1,'Same'] = stats_same
    df_stats.loc[1,'Changed'] = stats_changed
    df_stats.loc[1,'Added'] = stats_added
    df_stats.loc[1,'Removed'] = stats_removed
    print(df_stats.head(5))

    df_stats = df_stats.transpose()

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    timestr = time.strftime("%Y%m%d_%H%M%S")
    writer = pd.ExcelWriter(working_directory + 'ComparisonOutput_' + timestr + '.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    df_stats.to_excel(writer, sheet_name='Stats')
    df_new_diff.to_excel(writer, sheet_name='Diff')
    df_new.to_excel(writer, sheet_name='New')
    df_old.to_excel(writer, sheet_name='Old')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
    print("\nFile writer closed.\nPress any key to finish file output.")

    # Prevent command prompt closing
    input()


if __name__ == "__main__":
    main()