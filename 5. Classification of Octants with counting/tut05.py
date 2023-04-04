# Imported libraries required for programme (math, pandas, numpy, os, shutil) and cleared screen using cls command

from calendar import c
import math
import os
import pandas as pd
import numpy as np
import shutil
import pandas.io.formats.excel
from datetime import datetime
start_time = datetime.now()


os.system('cls')

# os.chdir() for changing directory for evaluation and smooth running, avoid confusions while writing code
os.chdir(r'C:\Users\Chandan\OneDrive\Documents\GitHub\2001CB21_2022\tut05')


# Function Defination
def octant_range_names(mod=5000):
    # Creating dictionary for octant name id mapping
    octant_name_id_mapping = {"1":"Internal outward interaction",
    "-1":"External outward interaction",
    "2":"External Ejection",
    "-2":"Internal Ejection",
    "3":"External inward interaction",
    "-3":"Internal inward interaction",
    "4":"Internal sweep",
    "-4":"External sweep"
    }
    # Try block if Input file not found
    try:
        # Copied content of input file into output file using shutil library and copyfile function
        # File is copied in target file, it is created if does not exist and overwrites if exists
        original = r'octant_input.xlsx'
        target = r'octant_output_ranking_excel.xlsx'
        shutil.copyfile(original, target)
    # Except statement if error occured in try block like input file does not exist or Correct Directory is not open
    except EnvironmentError:
        print("Oops! Error Occured. Please Check Input file present in Directory and Correct name given to it.")
        print("Or Check if any file is open anywhere, close it and retry")
        print("Correct input file name: octant_input.xlsx")
        exit()

    # Try block if output file do not created or exist
    try:
        # Created dataframe of output file using pandas library and read_excel() function 
        df= pd.read_excel('octant_output_ranking_excel.xlsx')

    # Except block if Output file is not found
    except:
        print("Output file not found.")
        exit()


    # Created size variable and u_mean, v_mean, w_mean variables and stored means of individual columns in that
    size_a=df['U'].size
    u_mean=(df['U']).mean()
    v_mean=(df['V']).mean()
    w_mean=(df['W']).mean()

    print("Please wait! Processing...\n")
    # Try Block if Output file not found then to except
    try:

        # Opening output file in append mode as write_obj
        with open('octant_output_ranking_excel.xlsx', 'a') as write_obj:
            #Created new blank columns for average velocities
            df["U avg"]= np.nan
            df["V avg"]= np.nan
            df["W avg"]= np.nan

            # Stored respective means at their position and column
            # .loc[] is used to access a group of rows and columns by labels
            df.loc[0,'U avg']=u_mean
            df.loc[0,'V avg']=v_mean
            df.loc[0,'W avg']=w_mean

            # Created empty columns for U' = U- U avg
            # np.nan fills the column with NaN values
            df["U' = U - U avg"]= np.nan
            # Filled empty columns with required values of difference between U and U avg
            for i in range(0, size_a):
                df.loc[i,"U' = U - U avg"]= df.loc[i,'U']-u_mean
            
            # Created empty columns for V' = V- V avg
            df["V' = V - V avg"]= np.nan
            # Filled empty columns with required values of difference between V and V avg
            for i in range(0, size_a):
                df.loc[i,"V' = V - V avg"]= df.loc[i,'V']-v_mean

            # Created empty columns for W' = W- W avg
            df["W' = W - W avg"]= np.nan
            # Filled empty columns with required values of difference between W and W avg
            for i in range(0, size_a):
                df.loc[i,"W' = W - W avg"]= df.loc[i,'W']-w_mean

            # Created an empty column for Octant Value storage
            df["Octant_value"]=np.nan
            
            # Logic for storing Octant Values
            '''Iterated through all rows and compared different permutations and assigned Octant Value'''
            # Running Loop for every row
            for i in range(0, size_a):
                # Comparison for selecting octant value
                if df.loc[i,"U' = U - U avg"]>=0 and df.loc[i,"V' = V - V avg"]>=0:
                    if df.loc[i,"W' = W - W avg"]>=0:
                        df.loc[i,"Octant_value"]=1
                    
                    else:
                        df.loc[i,"Octant_value"]=-1
                
                elif df.loc[i,"U' = U - U avg"]<0 and df.loc[i,"V' = V - V avg"]>=0:
                    if df.loc[i,"W' = W - W avg"]>=0:
                        df.loc[i,"Octant_value"]=2
                    
                    else:
                        df.loc[i,"Octant_value"]=-2
                
                elif df.loc[i,"U' = U - U avg"]<0 and df.loc[i,"V' = V - V avg"]<0:
                    if df.loc[i,"W' = W - W avg"]>=0:
                        df.loc[i,"Octant_value"]=3
                    
                    else:
                        df.loc[i,"Octant_value"]=-3
                
                else:
                    if df.loc[i,"W' = W - W avg"]>=0:
                        df.loc[i,"Octant_value"]=4
                    
                    else:
                        df.loc[i,"Octant_value"]=-4
                
            # Typecasting Octant values from float to int
            df["Octant_value"]=df["Octant_value"].astype('Int64')
                
            # Created column
            df[""]=np.nan

            # Just Created Good interphase for Excel file
            df.loc[0,"Octant ID"]="Overall Count"

            # Initialized count variables with zero   
            count_pos1=0
            count_neg1=0
            count_pos2=0
            count_neg2=0
            count_pos3=0
            count_neg3=0
            count_pos4=0
            count_neg4=0
                
            # Logic for counting octant values
            # Iterated through Octant Value column and compared with octant value and increased count by one        

            for i in range(0,size_a):
                if df.loc[i,"Octant_value"]==1:
                    count_pos1 +=1
                elif df.loc[i,"Octant_value"]==-1:
                    count_neg1 +=1
                elif df.loc[i,"Octant_value"]==2:
                    count_pos2 +=1
                elif df.loc[i,"Octant_value"]==-2:
                    count_neg2 +=1
                elif df.loc[i,"Octant_value"]==3:
                    count_pos3 +=1
                elif df.loc[i,"Octant_value"]==-3:
                    count_neg3 +=1
                elif df.loc[i,"Octant_value"]==4:
                    count_pos4 +=1
                elif df.loc[i, "Octant_value"]==-4:
                    count_neg4 +=1
            
            # Written values of counts at specified positions
            df.loc[0,"1"]=count_pos1
            df.loc[0,"-1"]=count_neg1
            df.loc[0,"2"]=count_pos2
            df.loc[0,"-2"]=count_neg2
            df.loc[0,"3"]=count_pos3
            df.loc[0,"-3"]=count_neg3        
            df.loc[0,"4"]=count_pos4   
            df.loc[0,"-4"]=count_neg4
            
            # Dictionary for mapping index to octant values
            util={
                1:1,
                2:-1,
                3:2,
                4:-2,
                5:3,
                6:-3,
                7:4,
                8:-4
            }
            # Logic: Creating a list for count values and creaating a copy of it. Sorting the copy of list.
            # Creating Dictionary of copy list and storing values as keys, indices as values
            ls = [count_pos1, count_neg1, count_pos2, count_neg2, count_pos3, count_neg3, count_pos4, count_neg4]
            ls_ = ls.copy()
            ls_.sort(reverse=True)
            ls_dict = {k: v for v, k in enumerate(ls_)}
            
            # Then Storing Ranks at specified positions
            for m in range(8):
                df.loc[0,f'Rank Octant {util[m+1]}']=ls_dict[ls[m]]+1
                if ls_dict[ls[m]]==0:
                    store=m
            # Creating Structure and storing Rank 1's Octant ID and its name
            df.loc[0,"Rank 1 Octant ID"] = util[store+1]
            df.loc[0, "Rank 1 Octant Name"] = octant_name_id_mapping[str(util[store+1])]
               
            # Beautify programme
            df.loc[1,""]="User Input"
            
            s = "mod "+ str(mod)
            df.loc[1,"Octant ID"]= s

            # Calculating no. of iterations for mod ranges using ceil function of Math library
            try:
                iter = math.ceil(size_a/ mod)
            # ZeroDivisionError for division by zero
            except ZeroDivisionError:
                print("Mod value cannot be zero! Please Change Mod value!!!")
                exit()

            # Creating a list for storing mod Ranks
            ranklist=[]
            # Iterating through each range and count the octant value of each ID
            for i in range(2, iter+2):
                # Writing range in file
                st=""
                if i==2:
                    x = min(mod*(i-1)-1,size_a)
                    st = "0000"+"-"+str(x)
                elif i== iter+1:
                    st= str(mod*(i-2))+"-"+str(size_a-1)
                else:
                    st = str(mod *(i-2))+"-"+str(mod*(i-1)-1)
                df.loc[i,"Octant ID"]= st

                # Initializing mod counts as zeroes
                mod_cnt_pos_one=0
                mod_cnt_neg_one=0
                mod_cnt_pos_two=0
                mod_cnt_neg_two=0
                mod_cnt_pos_three=0
                mod_cnt_neg_three=0
                mod_cnt_pos_four=0
                mod_cnt_neg_four=0

                # Handling current range for last call
                current_range= min(size_a, (i-1)*mod)

                # Iterating over range and incrementing count for respective mod count
                for j in range((i-2)*mod, current_range):
                    if df.loc[j, "Octant_value"]==1:
                        mod_cnt_pos_one+=1
                    elif df.loc[j, "Octant_value"]==-1:
                        mod_cnt_neg_one+=1
                    elif df.loc[j, "Octant_value"]==2:
                        mod_cnt_pos_two+=1
                    elif df.loc[j, "Octant_value"]==-2:
                        mod_cnt_neg_two+=1
                    elif df.loc[j, "Octant_value"]==3:
                        mod_cnt_pos_three+=1
                    elif df.loc[j, "Octant_value"]==-3:
                        mod_cnt_neg_three+=1
                    elif df.loc[j, "Octant_value"]==4:
                        mod_cnt_pos_four+=1
                    elif df.loc[j, "Octant_value"]==-4:
                        mod_cnt_neg_four+=1
                
                # writing mod counts at specified locations
                df.loc[i, "1"]=mod_cnt_pos_one
                df.loc[i, "-1"]=mod_cnt_neg_one
                df.loc[i, "2"]=mod_cnt_pos_two
                df.loc[i, "-2"]=mod_cnt_neg_two
                df.loc[i, "3"]=mod_cnt_pos_three
                df.loc[i, "-3"]=mod_cnt_neg_three
                df.loc[i, "4"]=mod_cnt_pos_four
                df.loc[i, "-4"]=mod_cnt_neg_four
                
                # Logic: Creating a list for count values and creaating a copy of it. Sorting the copy of list.
                # Creating Dictionary of copy list and storing values as keys, indices as values
                ls = [mod_cnt_pos_one, mod_cnt_neg_one, mod_cnt_pos_two, mod_cnt_neg_two, mod_cnt_pos_three, mod_cnt_neg_three, mod_cnt_pos_four, mod_cnt_neg_four]
                ls_ = ls.copy()
                ls_.sort(reverse=True)
                ls_dict = {k: v for v, k in enumerate(ls_)}
                # Then Storing Ranks at specified positions
                for m in range(8):
                    df.loc[i,f'Rank Octant {util[m+1]}']=ls_dict[ls[m]]+1
                    if ls_dict[ls[m]]==0:
                        store=m
                # Creating structure and storing Rank 1's Octant ID
                df.loc[i,"Rank 1 Octant ID"] = util[store+1]
                # Appending current Rank's octant to ranklist
                ranklist.append(util[store+1])
                # Storing current Octant's name
                df.loc[i, "Rank 1 Octant Name"] = octant_name_id_mapping[str(util[store+1])]
            
            # Creating structure 
            df.loc[iter+5,"1"]="Octant ID"
            df.loc[iter+5,"-1"]="Octant Name"
            df.loc[iter+5,"2"]="Count of Rank 1 Mod Values"
            # Iterating over loop and storing mod counts of each octant
            for n in range(iter+6, iter+14):
                df.loc[n,"1"]= (util[n-iter-5])
                df.loc[n,"-1"]=octant_name_id_mapping[str(util[n-iter-5])]
                df.loc[n,"2"]= ranklist.count(util[n-iter-5])

        # command for avoiding Bold header
        pandas.io.formats.excel.ExcelFormatter.header_style = None
        # try block for checking whether the file is open somewhere or not
        try:
            # Converting dataframe to excel file without index
            df.to_excel('octant_output_ranking_excel.xlsx',index=False)
            #After this file will be saved and automatically closed
        except:
            print("Output file is open somewhere. Please close and re-run the program")
            exit()
        

    # Except block if File not found or incorrect directory is opened
    except FileNotFoundError:
        print("Oops! Error Occured! File Not Found!")
        print("Please Check File present in correct directory with correct name")
        print("Correct File name is: octant_output_ranking_excel.xlsx")
        exit()
    

# You can change mod value here (if required)
mod=5000
# Function call 
octant_range_names(mod)


# Printing lines as Output is ready
print("Created required Output Excel file.")
print("Please Check 'octant_output_ranking_excel.xlsx' file")
end_time = datetime.now()
# Print Execution time
print('Duration of Program Execution: {}'.format(end_time - start_time))
exit()