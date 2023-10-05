import os
import shutil
from openpyxl import load_workbook



# wb = load_workbook('test.xlsx')
# ws = wb["Names in Badge"]

# Name in Badge (B)
# Affiliation (C)
# Role (D)

# ws.cell(row = 1, column=1).value


src_dir = os.getcwd() #get the current working dir
# print(src_dir)

# create a dir where we want to copy and rename
dest_dir = os.mkdir('Badges')
os.listdir()

dest_dir = src_dir+"/Badges"
src_file = os.path.join(src_dir, 'ATTENDEE.png')
shutil.copy(src_file,dest_dir) #copy the file to destination dir

dst_file = os.path.join(dest_dir,'ATTENDEE.png')
new_dst_file_name = os.path.join(dest_dir, 'name.png')

os.rename(dst_file, new_dst_file_name)#rename
os.chdir(dest_dir)


