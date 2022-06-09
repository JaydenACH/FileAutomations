import os
import openpyxl
import collections

# The function of this code is to loop through the folder, separate out the drawing numbers & revision numbers,
# and check for drawings with multiple revisions, keep the largest revision of that drawing numbers and
# write the values into a new Excel file.
# _____________________________________________________________________________________________________________

# Put in the directory that going to be the folder of all files
drawingdir = "C:\\Users\\Ang Chun Hang\\Documents\\SolidWork_Assy\\Offline Drawings"

# Change directory from directory of this python file to the directory above
os.chdir(drawingdir)

# Define some empty lists
drawingcol = []
revisioncol = []
drawinglist = []

# List out all files in folder and sub-folders
# Files with extension of .dwg or .DWG will be added into the drawinglist list
for roots, dirs, files in os.walk(drawingdir):
    for file in files:
        if file.endswith(".dwg") or file.endswith(".DWG"):
            if file not in drawinglist:
                drawinglist.append(file)

# Sort the drawing list in descending order (A-Z) and get the length of list
drawinglist.sort(reverse=False)
length = len(drawinglist)

# Loop through the whole list.
# Drop the "_rev#.dwg" and add into drawingcol list.
# Add # from rev into revisioncol list.
for i in range(length):
    drawingcol.append(drawinglist[i][:-9])
    revisioncol.append(int(drawinglist[i][-5]))

# Count how many revisions # for each drawing
mylist = collections.Counter(drawingcol)

# For each drawing (i) & number of drawings (j), if drawing is more than 1, check for larger revision number.
# Keep the larger revision number & drop the smaller revision number from list of drawing & revision.
# Length of list will -1 to prevent list looping out of index.
for i, j in mylist.items():
    if j > 2:
        d = 1
        while d < length:
            if drawingcol[d] == drawingcol[d - 1]:
                if revisioncol[d] > revisioncol[d - 1] or revisioncol[d] == revisioncol[d - 1]:
                    del drawingcol[d - 1]
                    del revisioncol[d - 1]
                    del drawinglist[d - 1]
                    length -= 1
            d += 1

# Create an Excel Workbook and activate the sheet, titled "Offline drawings revision"
# Save the new created Excel
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Offline drawings revision"
wb.save('Offline drawing revision.xlsx')

# Loop through new drawingcol and write drawingcol into column A and revisioncol into column B.
for i in range(len(drawingcol)):
    sheet.cell(row=i+1, column=1).value = drawingcol[i]
    sheet.cell(row=i+1, column=2).value = revisioncol[i]

# Save workbook after writing everything.
wb.save('Offline drawing revision.xlsx')
