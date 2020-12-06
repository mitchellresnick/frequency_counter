import csv

# ######################
# ALLISON EDIT SECTION

# The word in the header to exclude from analysis
headerKeyword = "Keywords"
# The column to analyze, this is zero indexed. Column A = 0, column B = 1, etc.
columnNum = 3
# #####################################

wordFrequency = {}


with open("C:/Users/atres/OneDrive/Desktop/Prospectus/Prospectus Writing/Book_10_Sample.csv", "r") as f:
    csvreader = csv.reader(f)
    for row in csvreader:
        if row[columnNum] != headerKeyword:
            for word in row[columnNum].split(' '):
                try:
                    wordFrequency[word] += 1
                except KeyError:
                    wordFrequency[word] = 1

with open("C:/Users/atres/OneDrive/Desktop/Prospectus/Prospectus Writing/Book_10_Sample_Results.csv", "w", newline='', encoding='utf-8') as o:
    csvwriter = csv.writer(o)
    csvwriter.writerow(["Keyword", "Count"])
    for (key, value) in wordFrequency.items():
        csvwriter.writerow([key, value])