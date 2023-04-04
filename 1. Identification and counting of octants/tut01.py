# Imported libraries required for programme (math , pandas, numpy, csv, os) and cleared screen using cls command

from calendar import c
import math
import os
import csv
import pandas as pd
import numpy as np
from datetime import datetime
start_time = datetime.now()

os.system('cls')


os.chdir(r'C:\Users\Chandan\OneDrive\Documents\GitHub\2001CB21_2022\tut01')



# Function defination
def octact_identification(mod=5000):
    # Opened input file in read mode and output file in write mode
    # Copied content of input file into output file using csv library and writerows function
    try:
        with open('octant_input.csv', 'r') as read_obj, open('octant_output.csv', 'w', newline='') as write_obj:
            csv_reader = csv.reader(read_obj)
            csv_writer =csv.writer(write_obj)
            csv_writer.writerows(csv_reader)


    # Except statement if error occured in try block like input file does not exist or Correct Directory is not open
    except EnvironmentError:
        print("Oops! Error Occured. Please Check Input file present in Directory and Correct name given to it.")
        print("Or Check if any file is open anywhere, close it and retry")
        print("Correct input file name: octant_input.csv")
        exit()

    # Created dataframe of output file using pandas library and read_csv() function 
    try:
        df= pd.read_csv('octant_output.csv')
    # Except block if Output file is not found
    except:
        print("Output file not found")
        exit()


    # Created size variable and u_mean, v_mean, w_mean variables and stored means of individual columns in that
    size_a=df['U'].size
    u_mean=(df['U']).mean()
    v_mean=(df['V']).mean()
    w_mean=(df['W']).mean()

    print("Please wait! Processing...")
    # Try Block if Output file not found then to except
    try:
        with open('octant_output.csv', 'a') as write_obj:
            
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
            # Filled empty columns with required values of difference
            for i in range(0, size_a):
                df.loc[i,"U' = U - U avg"]= df.loc[i,'U']-u_mean
            
            # Created empty columns for V' = V- V avg
            df["V' = V - V avg"]= np.nan
            # Filled empty columns with required values of difference
            for i in range(0, size_a):
                df.loc[i,"V' = V - V avg"]= df.loc[i,'V']-v_mean

            # Created empty columns for W' = W- W avg
            df["W' = W - W avg"]= np.nan
            # Filled empty columns with required values of difference
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
                
            # Created Empty column
            df[""]=np.nan

            # Just Created Good interphase for CSV file
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
            except ZeroDivisionError:
                print("Mod value cannot be zero.")
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
        # Saving changes to Output csv file  and closing file

        # try block for checking whether the file is open somewhere or not
        try:
            # Converting dataframe to CSV file without index
            df.to_csv('octant_output.csv',index=False)
            # After this file will be saved and automatically closed
        except:
            print("Output file is open somewhere. Please close it and re-run the program")
            exit()

    # Except block if File not found or incorrect directory is opened
    except FileNotFoundError:
        print("Oops! Error Occured! File Not Found!")
        print("Please Check File present in correct directory with correct name")
        print("Correct File name is: octant_output.csv")
        exit()

# You can change mod value here (if required)
mod=5000
# Function Call
octact_identification(mod)


# Printing lines as Output is ready
print("Created required Output CSV file.")
print("Please Check 'octant_output.csv' file")
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
exit()