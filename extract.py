import openpyxl
import re

#Open workbook and select the sheet
wb = openpyxl.load_workbook('Bioinformatics resources.xlsx')
ws = wb['Bioinformatics 4']

#Create a file that will contain hyperlinks
file = open("hyperlinks.txt", "w")

#Iterate trough the column that has hyperlinks and write them to a file
for x in range(1,303):
    hyperlink = (ws.cell(x, 7). value)
    hyperlink = str(hyperlink)
    file.write(hyperlink + "\n")
#Close the hyperlink file
file.close()

#Open hyperlinks file in read only mode
hyperlinks_file = open("hyperlinks.txt", 'r')
hyperlinks = hyperlinks_file.readlines()

#Identify the search pattern for filtering links and iterate trough the text file with the hyperlinks
pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
links = []
for line in hyperlinks:
    links += re.findall(pattern, line)

#Prints the final list of links
for link in links:
    print(link)
