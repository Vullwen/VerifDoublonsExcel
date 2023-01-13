import pandas as pd
import openpyxl
import time

# Open the input xlsx file
wb = openpyxl.load_workbook('20221227_filiere_electro_CARENE_daily.xlsx')

# Get the active sheet
sheet = wb.active

# Open the output txt file
with open('output.txt', 'w') as txt_file:
    # Iterate through the rows of the sheet
    for row in sheet.rows:
        # Get the cell in the "P" column
        statut = row[10]
        if statut.value != 'Supprimé':
            p_cell = row[15]
            if p_cell.value != "jobpath":
                #print(p_cell)
                # Split the cell value on the underscore
                if p_cell is not None:
                    p_cell_parts = p_cell.value.split('_')
                # Write the first part (before the underscore) to the txt file
                txt_file.write(p_cell_parts[0] + '\n')

print("Done!")

time.sleep(2)

# Open the output file in read mode
with open('output.txt', 'r') as output_file:
    #place each line in a list
    lines = output_file.readlines()
    #iterate through the list
    words = []
    for line in lines:
        #place on a list each word present more than once
        if lines.count(line) > 1:
            words.append(line)

#remove all the 2 last characters of each word
words = [word[:-1] for word in words]

time.sleep(2)

#create an dictionary with the words as keys and the number of occurences as values
d = {}
for word in words:
    if word in d:
        d[word] += 1
    else:
        d[word] = 1

#sort the dictionary
d = {k: v for k, v in sorted(d.items(), key=lambda item: item[1], reverse=False)}
#print the dictionary
print(d)

time.sleep(2)

# convert into dataframe
df = pd.DataFrame(data=d, index=[1])

#convert into excel
df.to_excel("iterations.xlsx", index=False)
print("Dictionary converted into excel...")


print("Fin d'exécution")