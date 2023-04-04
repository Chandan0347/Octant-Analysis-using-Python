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

os.chdir(r'C:\Users\Chandan\OneDrive\Documents\GitHub\2001CB21_2022\tut02')
# os.chdir() for changing directory for evaluation and smooth running, avoid confusions while writing code


# Function defination
def octant_transition_count(mod=5000):
    # Try block if Input file not found
    try:
        # Copied content of input file into output file using shutil library and copyfile function
        # File is copied in target file, it is created if does not exist and overwrites if exists
        original = r'input_octant_transition_identify.xlsx'
        target = r'output_octant_transition_identify.xlsx'
        shutil.copyfile(original, target)
    # Except statement if error occured in try block like input file does not exist or Correct Directory is not open
    except EnvironmentError:
        print("Oops! Error Occured. Please Check Input file present in Directory and Correct name given to it.")
        print("Or Check if any file is open anywhere, close it and retry")
        print("Correct input file name: input_octant_transition_identify.xlsx")
        exit()

    # Try block if output file do not created or exist
    try:
        # Created dataframe of output file using pandas library and read_excel() function 
        df= pd.read_excel('output_octant_transition_identify.xlsx')

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
        with open('output_octant_transition_identify.xlsx', 'a') as write_obj:
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

            # Initializing variables for Verification of counts
            pos1=0
            neg1=0
            pos2=0
            neg2=0
            pos3=0
            neg3=0
            pos4=0
            neg4=0

            # Iterating over the count of ranges and summing up
            for i in range(iter):
                pos1+= df.loc[i+2,"1"]
                neg1+= df.loc[i+2,"-1"]
                pos2+= df.loc[i+2,"2"]
                neg2+= df.loc[i+2,"-2"]
                pos3+= df.loc[i+2,"3"]
                neg3+= df.loc[i+2,"-3"]
                pos4+= df.loc[i+2,"4"]
                neg4+= df.loc[i+2,"-4"]

            # Placing Values at different locations
            df.loc[iter+2,"Octant ID"]="Verified"
            df.loc[iter+2,"1"]=pos1
            df.loc[iter+2,"-1"]=neg1
            df.loc[iter+2,"2"]=pos2
            df.loc[iter+2,"-2"]=neg2
            df.loc[iter+2,"3"]=pos3
            df.loc[iter+2,"-3"]=neg3
            df.loc[iter+2,"4"]=pos4
            df.loc[iter+2,"-4"]=neg4

            #Creating Structure for Overall Transition count Table

            df.loc[iter+5, "Octant ID"]="Overall Transition Count"
            df.loc[iter+7,"Octant ID"]="Count"
            df.loc[iter+8,""]="From"
            df.loc[iter+6,"1"]="To"
            df.loc[iter+8,"Octant ID"]="+1"
            df.loc[iter+9,"Octant ID"]="-1"
            df.loc[iter+10,"Octant ID"]="+2"
            df.loc[iter+11,"Octant ID"]="-2"
            df.loc[iter+12,"Octant ID"]="+3"
            df.loc[iter+13,"Octant ID"]="-3"
            df.loc[iter+14,"Octant ID"]="+4"
            df.loc[iter+15,"Octant ID"]="-4"

            df.loc[iter+7,"1"]="+1"
            df.loc[iter+7,"-1"]="-1"
            df.loc[iter+7,"2"]="+2"
            df.loc[iter+7,"-2"]="-2"
            df.loc[iter+7,"3"]="+3"
            df.loc[iter+7,"-3"]="-3"
            df.loc[iter+7,"4"]="+4"
            df.loc[iter+7,"-4"]="-4"

            # Initializing every cell in table as 0
            for i in range(8):
                df.loc[iter+8+i,"1"]=0
                df.loc[iter+8+i,"-1"]=0
                df.loc[iter+8+i,"2"]=0
                df.loc[iter+8+i,"-2"]=0
                df.loc[iter+8+i,"3"]=0
                df.loc[iter+8+i,"-3"]=0
                df.loc[iter+8+i,"4"]=0
                df.loc[iter+8+i,"-4"]=0

            
            # Counting and filling the table of overall transition count
            for i in range(0,size_a-1):
                x= df.loc[i,"Octant_value"]
                y=df.loc[i+1,"Octant_value"]
                z=0
                if x>0:
                    z= x*2 -1
                elif x<0:
                    z= -2 * x
                
                df.loc[iter+7+z,f'{y}']=df.loc[iter+7+z, f'{y}']+1
            
            # Counting the mod ranges using this for loop
            for i in range(iter):
                # Creating structures for mod ranges count
                df.loc[iter+19+13*i, "Octant ID"]="Mod Transition Count"
                df.loc[iter+21+13*i,"Octant ID"]="Count"
                df.loc[iter+22+13*i,""]="From"
                df.loc[iter+20+13*i,"1"]="To"
                df.loc[iter+22+13*i,"Octant ID"]="+1"
                df.loc[iter+23+13*i,"Octant ID"]="-1"
                df.loc[iter+24+13*i,"Octant ID"]="+2"
                df.loc[iter+25+13*i,"Octant ID"]="-2"
                df.loc[iter+26+13*i,"Octant ID"]="+3"
                df.loc[iter+27+13*i,"Octant ID"]="-3"
                df.loc[iter+28+13*i,"Octant ID"]="+4"
                df.loc[iter+29+13*i,"Octant ID"]="-4"

                df.loc[iter+21+13*i,"1"]="+1"
                df.loc[iter+21+13*i,"-1"]="-1"
                df.loc[iter+21+13*i,"2"]="+2"
                df.loc[iter+21+13*i,"-2"]="-2"
                df.loc[iter+21+13*i,"3"]="+3"
                df.loc[iter+21+13*i,"-3"]="-3"
                df.loc[iter+21+13*i,"4"]="+4"
                df.loc[iter+21+13*i,"-4"]="-4"

                # Initializing every cell of table to zero
                for j in range(8):
                    df.loc[iter+22+j+13*i,"1"]=0
                    df.loc[iter+22+j+13*i,"-1"]=0
                    df.loc[iter+22+j+13*i,"2"]=0
                    df.loc[iter+22+j+13*i,"-2"]=0
                    df.loc[iter+22+j+13*i,"3"]=0
                    df.loc[iter+22+j+13*i,"-3"]=0
                    df.loc[iter+22+j+13*i,"4"]=0
                    df.loc[iter+22+j+13*i,"-4"]=0

                # Initializing start and end variables
                start = i * mod
                end = ((i+1) * mod) if ((i+1) * mod) < size_a else size_a-1
                # Creating string for mod ranges and storing at correct location
                if i==0:
                    x = min(mod*(i+1),size_a)
                    s = "0000"+"-"+str(x-1)
                elif i== iter-1:
                    s= str(start)+"-"+str(size_a-1)
                else:
                    s = str(start)+"-"+str(end-1)
                df.loc[iter+20+13*i,"Octant ID"]= s
                
                #Storing count of transition in cell of table after iterating through range 
                for k in range(start,end):
                    x= df.loc[k,"Octant_value"]
                    y= df.loc[k+1,"Octant_value"]
                    z=0
                    if x>0:
                        z= x*2 -1
                    elif x<0:
                        z= -2 * x
                    
                    df.loc[iter+21+13*i+z,f'{y}']=df.loc[iter+21+z+13*i, f'{y}']+1
        # command for avoiding Bold header
        pandas.io.formats.excel.ExcelFormatter.header_style = None
        # try block for checking whether the file is open somewhere or not
        try:
            # Converting dataframe to excel file without index
            df.to_excel('output_octant_transition_identify.xlsx',index=False)
            #After this file will be saved and automatically closed
        except:
            print("Output file is open somewhere. Please close and re-run the program")
            exit()
        

    # Except block if File not found or incorrect directory is opened
    except FileNotFoundError:
        print("Oops! Error Occured! File Not Found!")
        print("Please Check File present in correct directory with correct name")
        print("Correct File name is: output_octant_transition_identify.xlsx")
        exit()
        
# You can change mod value here
mod=5000
# Function call
octant_transition_count(mod)
# Printing lines as Output is ready
print("Created required Output Excel file.")
print("Please Check 'output_octant_transition_identify.xlsx' file")
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
exit()