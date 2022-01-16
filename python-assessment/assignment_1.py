import fitz
import pandas as pd

#Add PDF Path
path = '...keppel.pdf'

#open pdf document
doc = fitz.open(path) 

#ADD PREFERRED PAGE NUMBER IN THE PARAMETER
#NUMBER STARTS WITH 0
#Change page number here
page = doc.load_page(11) 

#get text blocks from PDF
block = page.get_text('blocks',sort = True)

# converting block to dataframe and sorting values by x,y coordinates and block number
df = pd.DataFrame(block)
df.sort_values([0,1,3,5], ascending=[True,True,True,True], inplace=True) 
df[7] = df[0] - df[0].shift() #adds a column to the dataframe where it computes values of previous x coordinate to current

#initializing empty list to store per column
A = []
B = []
C = []
count = 0 #a counter for what column to use/store

for index, row in df.iterrows():
    if row[7] > 100 and (row[7] != 0 or row[7] != NaN): #check difference if it changes
        count +=1
        if count == 0:
            A.append(row)
        elif count == 1:
            B.append(row)
        else:
            C.append(row)
    else:
        if count == 0:
            A.append(row)
        elif count == 1:
            B.append(row)
        else:
            C.append(row)
    

# converting list A,B,C to dataframe and sort values by Y coordinates

df1 = pd.DataFrame(A, index=None) #Column 1
df2 = pd.DataFrame(B, index=None) #Column 2
df3 = pd.DataFrame(C, index=None) #Column 3
df1.sort_values([1], ascending=[True], inplace=True)
df2.sort_values([1], ascending=[True], inplace=True)
df3.sort_values([1], ascending=[True], inplace=True)



# export to_excel by pandas using xlsxWriter engine.
writer = pd.ExcelWriter('assignment_1.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
df1.to_excel(writer, sheet_name='Column 1', columns=[4],header=False, index=False)
df2.to_excel(writer, sheet_name='Column 2', columns=[4],header=False, index=False)
df3.to_excel(writer, sheet_name='Column 3', columns=[4],header=False, index=False)

#close the writer engine and save excel file
#file is saved in the same folder
writer.save()