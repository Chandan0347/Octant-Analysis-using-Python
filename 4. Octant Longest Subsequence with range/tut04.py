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
os.chdir(r'C:\Users\Chandan\OneDrive\Documents\GitHub\2001CB21_2022\tut04')

# Function defination
def octant_longest_subsequence_count_with_range():
    # Try block if Input file not found
    try:
        # Copied content of input file into output file using shutil library and copyfile function
        # File is copied in target file, it is created if does not exist and overwrites if exists
        original = r'input_octant_longest_subsequence_with_range.xlsx'
        target = r'output_octant_longest_subsequence_with_range.xlsx'
        shutil.copyfile(original, target)
    # Except statement if error occured in try block like input file does not exist or Correct Directory is not open
    except EnvironmentError:
        print("Oops! Error Occured. Please Check Input file present in Directory and Correct name given to it.")
        print("Or Check if any file is open anywhere, close it and retry")
        print("Correct input file name: input_octant_longest_subsequence_with_range.xlsx")
        exit()

    # Try block if output file did not created or exist
    try:
        # Created dataframe of output file using pandas library and read_excel() function 
        df = pd.read_excel('output_octant_longest_subsequence_with_range.xlsx')

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
        with open('output_octant_longest_subsequence_with_range.xlsx', 'a') as write_obj:
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
                
            # Created rows for different octants
            df[""]=np.nan
            df["Octant"]=np.nan
            df.loc[0, "Octant"]="+1"
            df.loc[1, "Octant"]="-1"
            df.loc[2, "Octant"]="+2"
            df.loc[3, "Octant"]="-2"
            df.loc[4, "Octant"]="+3"
            df.loc[5, "Octant"]="-3"
            df.loc[6, "Octant"]="+4"
            df.loc[7, "Octant"]="-4"

            # Created columns Longest subsequence length and count
            df["Longest Subsequence Length"]=np.nan
            df["Count"]=np.nan

            # Initialized variables for max length of subsequence of each octant
            max_length_p1=0
            max_length_n1=0
            max_length_p2=0
            max_length_n2=0
            max_length_p3=0
            max_length_n3=0
            max_length_p4=0
            max_length_n4=0

            # Initialized Count variables for each octant value
            count_p1 = 0
            count_n1 = 0
            count_p2 = 0
            count_n2 = 0
            count_p3 = 0
            count_n3 = 0
            count_p4 = 0
            count_n4 = 0
            # Initialized time range variables of octant as list to store j
            time_rangesp1 = []
            time_rangesn1 = []
            time_rangesp2 = []
            time_rangesn2 = []
            time_rangesp3 = []
            time_rangesn3 = []
            time_rangesp4 = []
            time_rangesn4 = []


            # Iterated through overall range of octant columns
            for j in range(size_a):
                # Initialized curr_length variable as 0 for keep track of length of current octant value subsequence
                curr_length = 0
                # Variable x for storing Octant Value
                x = df.loc[j,"Octant_value"]
                # Current index as idx
                idx = j

                # Comparing current octant value with each 8 octant values and incrementing corresponding 
                # max subsequence length variable and count of that max subsequence
                # For 1
                if x==1:
                    # Increasing current length by 1
                    curr_length +=1
                    # Checking if current length is greater than max subsequence length till now then count of that octant will become 1
                    if curr_length>max_length_p1:
                        count_p1=1
                        # Clearing time range if new max length found and appending that index to index range
                        time_rangesp1.clear()
                        time_rangesp1.append([df.loc[idx, "Time"], df.loc[idx , "Time"]])
                    # If current length is equal to the maximum length then increase the count of maximum subsequence length
                    elif curr_length==max_length_p1:
                        count_p1+=1
                        # Appending the list for from time to time as list to time range for current octant
                        time_rangesp1.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    # storing maximum subsequence length by comparing it with current length
                    max_length_p1 = max(max_length_p1, curr_length)
                
                # For -1
                if x==-1:
                    curr_length +=1
                    if curr_length>max_length_n1:
                        count_n1=1
                        time_rangesn1.clear()
                        time_rangesn1.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_n1:
                        count_n1+=1
                        time_rangesn1.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_n1 = max(max_length_n1, curr_length)

                # For 2
                if x==2:
                    curr_length +=1
                    if curr_length>max_length_p2:
                        count_p2=1
                        time_rangesp2.clear()
                        time_rangesp2.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_p2:
                        count_p2+=1
                        time_rangesp2.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_p2 = max(max_length_p2, curr_length)
                # For -2
                if x==-2:
                    curr_length +=1
                    if curr_length>max_length_n2:
                        count_n2=1
                        time_rangesn2.clear()
                        time_rangesn2.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_n2:
                        count_n2+=1
                        time_rangesn2.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_n2 = max(max_length_n2, curr_length)
                # For 3
                if x==3:
                    curr_length +=1
                    if curr_length>max_length_p3:
                        count_p3=1
                        time_rangesp3.clear()
                        time_rangesp3.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_p3:
                        count_p3+=1
                        time_rangesp3.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_p3 = max(max_length_p3, curr_length)
                # For -3
                if x==-3:
                    curr_length +=1
                    if curr_length>max_length_n3:
                        count_n3=1
                        time_rangesn3.clear()
                        time_rangesn3.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_n3:
                        count_n3+=1
                        time_rangesn3.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_n3 = max(max_length_n3, curr_length)
                # For 4
                if x==4:
                    curr_length +=1
                    if curr_length>max_length_p4:
                        count_p4=1
                        time_rangesp4.clear()
                        time_rangesp4.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_p4:
                        count_p4+=1
                        time_rangesp4.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_p4 = max(max_length_p4, curr_length)
                # For -4
                if x==-4:
                    curr_length +=1
                    if curr_length>max_length_n4:
                        count_n4=1
                        time_rangesn4.clear()
                        time_rangesn4.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    elif curr_length==max_length_n4:
                        count_n4+=1
                        time_rangesn4.append([df.loc[idx, "Time"], df.loc[idx, "Time"]])
                    max_length_n4 = max(max_length_n4, curr_length)

                # If j becomes last index, exited the loop by break statement because its work is done
                if j== size_a -1:
                    break

                # Running while loop till the current Octant value is equal to next Octant value
                try:
                    while df.loc[j,"Octant_value"]==df.loc[j+1,"Octant_value"]:
                        # For 1
                        if x==1:
                            # Increasing current length by 1
                            curr_length +=1
                            # Checking if current length is greater than max subsequence length till now then count of that octant will become 1
                            if curr_length>max_length_p1:
                                count_p1=1
                                # Clearing time range if new max length found and appending that index to index range
                                time_rangesp1.clear()
                                time_rangesp1.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            # If current length is equal to the maximum length then increase the count of maximum subsequence length
                            elif curr_length==max_length_p1:
                                count_p1+=1
                                # Appending the list for from time to time as list to time range for current octant
                                time_rangesp1.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            # storing maximum subsequence length by comparing it with current length
                            max_length_p1 = max(max_length_p1, curr_length)
                        # For -1
                        if x==-1:
                            curr_length +=1
                            if curr_length>max_length_n1:
                                count_n1=1
                                time_rangesn1.clear()
                                time_rangesn1.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_n1:
                                count_n1+=1
                                time_rangesn1.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_n1 = max(max_length_n1, curr_length)
                        # For 2
                        if x==2:
                            curr_length +=1
                            if curr_length>max_length_p2:
                                count_p2=1
                                time_rangesp2.clear()
                                time_rangesp2.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_p2:
                                count_p2+=1
                                time_rangesp2.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_p2 = max(max_length_p2, curr_length)
                        # For -2
                        if x==-2:
                            curr_length +=1
                            if curr_length>max_length_n2:
                                count_n2=1
                                time_rangesn2.clear()
                                time_rangesn2.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_n2:
                                count_n2+=1
                                time_rangesn2.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_n2 = max(max_length_n2, curr_length)
                        # For 3
                        if x==3:
                            curr_length +=1
                            if curr_length>max_length_p3:
                                count_p3=1
                                time_rangesp3.clear()
                                time_rangesp3.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_p3:
                                count_p3+=1
                                time_rangesp3.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_p3 = max(max_length_p3, curr_length)
                        # For -3
                        if x==-3:
                            curr_length +=1
                            if curr_length>max_length_n3:
                                count_n3=1
                                time_rangesn3.clear()
                                time_rangesn3.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_n3:
                                count_n3+=1
                                time_rangesn3.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_n3 = max(max_length_n3, curr_length)
                        # For 4
                        if x==4:
                            curr_length +=1
                            if curr_length>max_length_p4:
                                count_p4=1
                                time_rangesp4.clear()
                                time_rangesp4.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_p4:
                                count_p4+=1
                                time_rangesp4.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_p4 = max(max_length_p4, curr_length)
                        # For -4
                        if x==-4:
                            curr_length +=1
                            if curr_length>max_length_n4:
                                count_n4=1
                                time_rangesn4.clear()
                                time_rangesn4.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            elif curr_length==max_length_n4:
                                count_n4+=1
                                time_rangesn4.append([df.loc[idx, "Time"], df.loc[j+1, "Time"]])
                            max_length_n4 = max(max_length_n4, curr_length)
                        
                        # Incrementing j by 1
                        j = j+1

                        # If j equals to last index after incrementing after while loop then we should compare
                        # it with current octant or else it will lead to wrong output
                        if j == size_a-1:
                            # For 1
                            if x==1:
                                # Increasing current length by 1
                                curr_length +=1
                                # Checking if current length is greater than max subsequence length till now then count of that octant will become 1
                                if curr_length>max_length_p1:
                                    count_p1=1
                                    # Clearing time range if new max length found and appending that index to index range
                                    time_rangesp1.clear()
                                    time_rangesp1.append([df.loc[idx, "Time"], df.loc[j , "Time"]])
                                # If current length is equal to the maximum length then increase the count of maximum subsequence length
                                elif curr_length==max_length_p1:
                                    count_p1+=1
                                    # Appending the list for from time to time as list to time range for current octant
                                    time_rangesp1.append([df.loc[idx, "Time"], df.loc[j , "Time"]])
                                # storing maximum subsequence length by comparing it with current length
                                max_length_p1 = max(max_length_p1, curr_length)
                            # For -1
                            if x==-1:
                                curr_length +=1
                                if curr_length>max_length_n1:
                                    count_n1=1
                                    time_rangesn1.clear()
                                    time_rangesn1.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_n1:
                                    count_n1+=1
                                    time_rangesn1.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_n1 = max(max_length_n1, curr_length)
                            # For 2
                            if x==2:
                                curr_length +=1
                                if curr_length>max_length_p2:
                                    count_p2=1
                                    time_rangesp2.clear()
                                    time_rangesp2.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_p2:
                                    count_p2+=1
                                    time_rangesp2.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_p2 = max(max_length_p2, curr_length)
                            # For -2
                            if x==-2:
                                curr_length +=1
                                if curr_length>max_length_n2:
                                    count_n2=1
                                    time_rangesn2.clear()
                                    time_rangesn2.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_n2:
                                    count_n2+=1
                                    time_rangesn2.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_n2 = max(max_length_n2, curr_length)
                            # For 3
                            if x==3:
                                curr_length +=1
                                if curr_length>max_length_p3:
                                    count_p3=1
                                    time_rangesp3.clear()
                                    time_rangesp3.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_p3:
                                    count_p3+=1
                                    time_rangesp3.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_p3 = max(max_length_p3, curr_length)
                            # For -3
                            if x==-3:
                                curr_length +=1
                                if curr_length>max_length_n3:
                                    count_n3=1
                                    time_rangesn3.clear()
                                    time_rangesn3.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_n3:
                                    count_n3+=1
                                    time_rangesn3.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_n3 = max(max_length_n3, curr_length)
                            # For 4
                            if x==4:
                                curr_length +=1
                                if curr_length>max_length_p4:
                                    count_p4=1
                                    time_rangesp4.clear()
                                    time_rangesp4.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_p4:
                                    count_p4+=1
                                    time_rangesp4.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_p4 = max(max_length_p4, curr_length)
                            # For -4
                            if x==-4:
                                curr_length +=1
                                if curr_length>max_length_n4:
                                    count_n4=1
                                    time_rangesn4.clear()
                                    time_rangesn4.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                elif curr_length==max_length_n4:
                                    count_n4+=1
                                    time_rangesn4.append([df.loc[idx, "Time"], df.loc[j, "Time"]])
                                max_length_n4 = max(max_length_n4, curr_length)
                            # Breaking the for loop for last index
                            break
                except:
                    print("Programme Failed :(")
                    exit()
            
                
            # Finally storing max subsequence length of octant values at specified positions
            df.loc[0,"Longest Subsequence Length"]= max_length_p1
            df.loc[1,"Longest Subsequence Length"]= max_length_n1
            df.loc[2,"Longest Subsequence Length"]= max_length_p2
            df.loc[3,"Longest Subsequence Length"]= max_length_n2
            df.loc[4,"Longest Subsequence Length"]= max_length_p3
            df.loc[5,"Longest Subsequence Length"]= max_length_n3
            df.loc[6,"Longest Subsequence Length"]= max_length_p4
            df.loc[7,"Longest Subsequence Length"]= max_length_n4

            # Storing Count of each maximum length of octant values at specified locations in Count column
            df.loc[0,"Count"]= count_p1
            df.loc[1,"Count"]= count_n1
            df.loc[2,"Count"]= count_p2
            df.loc[3,"Count"]= count_n2
            df.loc[4,"Count"]= count_p3
            df.loc[5,"Count"]= count_n3
            df.loc[6,"Count"]= count_p4
            df.loc[7,"Count"]= count_n4

            # Creating columns for Storing time ranges
            df.insert(15,"",np.nan,True)
            df["_Octant_"]=np.nan
            df["Longest_Subsequence_Length"]=np.nan
            df["Count_"]=np.nan
            
            # Creating dictionary for Octant value
            dict_octant ={
                0 : "+1",
                1 : "-1",
                2 : "+2",
                3 : "-2",
                4 : "+3",
                5 : "-3",
                6 : "+4",
                7 : "-4",
            }
            # Creating dictionary for counts of longest subsequence length of octant values
            dict_count = {
                0: count_p1,
                1: count_n1,
                2: count_p2,
                3: count_n2,
                4: count_p3,
                5: count_n3,
                6: count_p4,
                7: count_n4,
            }
            # Creating dictionary for maximum length of subsequence of octant values
            dict_length = {
                0: max_length_p1,
                1: max_length_n1,
                2: max_length_p2,
                3: max_length_n2,
                4: max_length_p3,
                5: max_length_n3,
                6: max_length_p4,
                7: max_length_n4
            }
            # Creating dictionary for time ranges of octant values
            dict_time={
                0: time_rangesp1,
                1: time_rangesn1,
                2: time_rangesp2,
                3: time_rangesn2,
                4: time_rangesp3,
                5: time_rangesn3,
                6: time_rangesp4,
                7: time_rangesn4,

            }
            # Variable
            k =0

            # For loop for writing diffrent values of time ranges in dataframe
            for i in range(8):
                # Structure of output
                df.loc[k,"_Octant_"]= dict_octant[i]
                df.loc[k,"Longest_Subsequence_Length"]= dict_length[i]
                df.loc[k,"Count_"]= dict_count[i]
                df.loc[k+1,"_Octant_"]= "Time"
                df.loc[k+1,"Longest_Subsequence_Length"]= "From"
                df.loc[k+1,"Count_"]= "To"
                #Variable
                l=0
                # For loop for writing time ranges at specific position
                for item in dict_time[i]:
                    df.loc[k+2+l,"Longest_Subsequence_Length"]=item[0]
                    df.loc[k+2+l,"Count_"]=item[1]
                    l=l+1
                # Incrementing k+2 by current count of current octant
                k= k+2 + dict_count[i]

        # Command for avoiding Bold header
        pandas.io.formats.excel.ExcelFormatter.header_style = None
        # try block for checking whether the file is open somewhere or not
        try:
            # Converting dataframe to excel file without index
            df.to_excel('output_octant_longest_subsequence_with_range.xlsx',index=False)
            #After this file will be saved and automatically closed
        except:
            print("Output file is open somewhere. Please close and re-run the program")
            exit()
        

    # Except block if File not found or incorrect directory is opened
    except FileNotFoundError:
        print("Oops! Error Occured! File Not Found!")
        print("Please Check File present in correct directory with correct name")
        print("Correct File name is: output_octant_longest_subsequence_with_range.xlsx")
        exit()
    
# Function call
octant_longest_subsequence_count_with_range()
# Printing lines as Output is ready
print("Created required Output Excel file.")
print("Please Check 'output_octant_longest_subsequence_with_range.xlsx' file")
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
exit()

