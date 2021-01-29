# Tool to compare table documents

# import modules
import pandas as pd
# also required for excel files: xlrd AND openpyxl
# Don't wrap repr(DataFrame) across additional lines
#pd.set_option("display.expand_frame_repr", False)


def fcn_compare(val1,val2):
    if val1 == val2:
        return True
    else:
        return False


# welcoming
print("Welcome to CompareX")
print("Which tables do you want to read?")

# user input, which tables to read

# ...later


# open tables and read to variables
try:
    df_compare_a = pd.read_excel("CompareA.xlsx","Tabelle1")
    print("Table A read successfully.")
except:
    print("An error occured, reading comparison file A.")

try:
    df_compare_b = pd.read_excel("CompareB.xlsx","Tabelle1")
    print("Table B read successfully.")
except:
    print("An error occured, reading comparison file B.")

merged_df = pd.merge(df_compare_a,df_compare_b,on=['Nr'],how="outer",suffixes=('_old','_new'))

#merged_df.to_csv("out.csv")


a_rows = df_compare_a.shape[0]
a_cols = df_compare_a.shape[1]
b_rows = df_compare_b.shape[0]
b_cols = df_compare_b.shape[1]

df_result = 0



# count through rows
for c_row in range(0,merged_df.shape[0]):

    # initialize start columns
    df_merged_colZero_old = 1
    df_merged_colZero_new = df_merged_colZero_old + a_cols-1

    for c_col in range(1,a_cols):
        cell_old = merged_df.iat[c_row,c_col]
        cell_new = merged_df.iat[c_row,c_col+a_cols-1]
        
        cmp_result = fcn_compare(cell_old,cell_new)

        df_result.iat[c_cow,c_col] = cmp_result


print(df_result)
    # count through cols
    #print(merged_df.iat[c_row,df_merged_colZero_old])
    #print(merged_df.iat[c_row,df_merged_colZero_new])

    