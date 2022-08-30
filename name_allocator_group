# -*- coding: utf-8 -*-
"""
Important notes:

This code will generate a folder to store name-data excel files on Desktop, 
the folder is named as 'Camp file', there should be two excel files and one text
file.

To use this code, need to install the following modules ---
pandas + openpyxl + os, 
they are easy to install, installations can be found in Google.

The code takes the orgianl excel file path, which contains the raw name data, as input.

Path has to be in a certain form: (remember the 'r' in the front)
    file = r'excel path'
    


"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
import os
import random
########################### Useful functions##################################
def gender_trans(cell):
    '''
    Translate gender from Chinese to English
    Parameters
    
    ----------
    cell : Str
        String read from the cell in Gender column

    Returns
    -------
    str
        Gender English Abbr. 

    '''
    if cell == '男':
        return 'M'
    if cell == '女':
        return 'F'
    
def person_trans(cell):
    '''
    Translate personalities from chinese to English. There are four in total.
    情感外向型, 情感内向型, 理智外向型, 理智内向型
    
    Parameters
    ----------
    cell : Str
        String read from the cell in Personality column

    Returns
    -------
    str
        Personality English Abbr.

    '''
    if cell == '情感外向型':
        return 'EE'
    if cell == '情感内向型':
        return 'EI'
    if cell == '理智外向型':
        return 'RE'
    if cell == '理智内向型':
        return 'RI'
    else:
        return '不确定'
##############################################################################
###########################  Editions from here  #############################

'''
Headers and sheet names can be edited in 'Preparing the dataset' and 
'Sheet headers' sections.
'''

# Indicate excel sheet 'path'!!!  Remeber the 'r'
file = r'C:\Users\work\Desktop\test0.xlsx'

# Sheet format, 0 = headers on, None = No header 
header_ = 0   

# Constrain on number of servings (will be used to highlight the name in excel)
max_serve = 4   

# Number of groups (will be used to allocate members to groups)
num_group = 12

# Name of the camp
camp_name = '2021YEcamp'
#########################        End here        #############################

##############################################################################

# Creating a folder for excel and text files on Desktop
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
newpath = desktop + '\Camp file' 
if not os.path.exists(newpath):
    os.makedirs(newpath)
    
# Save a copy to a different file
newbook = r'%s\camp.xlsx'%(newpath)    # excel file for general classification
serbook = r'%s\serving.xlsx'%(newpath) # excel file for serving teams
newtxt = r'%s\stats.txt'%(newpath)     # text file for statistics

########################  Preparing the dataset  #############################

#---------------- This part is for general classification -------------------#
# Read the sheet, copy the excel flie to a new folder at desktop
# and prepare several sheets for classifications
df0 = pd.read_excel(file, header=header_)       # Read the original data
df0.to_excel(newbook)       # Duplicate the orginal copy
df = load_workbook(newbook)
sheet_name0 = ['邮箱跟踪','给义工','分组','备注以及住宿要求','食物过敏','营费补助']
for i in sheet_name0:                       # Making sheets for general file
    df.create_sheet(i)
df.save(newbook)

# Reading data
df = pd.read_excel(newbook, header=header_)  
# Preparing sheets for writing data
dff = load_workbook(newbook)
ws0 = dff.worksheets[0]         #'名单'
Email = dff.worksheets[1]       #'邮箱跟踪'
Room = dff.worksheets[2]        #'给义工'       \\\\\\\\\\\\\
Group = dff.worksheets[3]       #'分组'         \\\\\\\\\\\\\
Request = dff.worksheets[4]     #'备注以及住宿要求'
Allergy = dff.worksheets[5]     #'食物过敏'
Scholar = dff.worksheets[6]     #'营费补助'
#-----------------------------------------------------------------------------#

#--------------------- This part is for serving teams ------------------------#
df0.to_excel(serbook)                         # Duplicate the orginal copy
sdf = load_workbook(serbook)
sheet_name1 = ['小组长备选','带领敬拜','敬拜乐手','PA音响控制','后勤清理','分享见证','祷告团队','Revisit']
for i in sheet_name1:                         # Making sheets for serving file
    sdf.create_sheet(i)
sdf.save(serbook)

# Reading data
sdf = pd.read_excel(serbook, header=header_)  
# Preparing sheets for writing data
sdff = load_workbook(serbook)
sws0 = sdff.worksheets[0]         #'名单'
sGleader = sdff.worksheets[1]     #'小组长备选'
sWorship = sdff.worksheets[2]     #'带领敬拜'
sInstru = sdff.worksheets[3]      #'敬拜乐手'
sPA = sdff.worksheets[4]          #'PA音响控制'        
sClean = sdff.worksheets[5]       #'后勤清理'
sTestm = sdff.worksheets[6]       #'分享见证'
sPrayer = sdff.worksheets[7]      #'祷告团队'
sRevisit = sdff.worksheets[8]     #'遗漏'
#-----------------------------------------------------------------------------#
###############################################################################

############################### Sheet headers #################################

# Preparing the headers for each sheet
columns = df.columns              # list of columns' name from orignal data set
num_row = len(df[columns[0]])     # number of rows in data

size = ['XS','S','M','L','XL','2XXL']
email = ['中文姓名','英文名（或者拼音）','邮箱']
request = ['ID','中文姓名', '性别','备注']
allergy = ['ID','中文姓名', '性别','食物过敏']
scholar = ['ID','中文姓名', '性别','申请营费补助原因','申请营费补助费用']

serve = ['小组长（需要提前抵达营会进行小组长培训）','带领敬拜','敬拜乐手（请在others标明乐器）','PA音响控制','后勤清理','分享见证','祷告团队']
serve_sheet = [sGleader,sWorship,sInstru,sPA,sClean,sTestm,sPrayer,sRevisit]


test_serve = ['小组长','带领敬拜','敬拜乐手（请在others标明乐器）','PA音响控制','后勤清理','分享见证','祷告团队']

time_believe = ['大于5年','3-5年','2年','1年','少于1年','还未信主']
group_header = ['ID','中文姓名','性别','性格倾向']
###############################################################################

######################### Global variables and values #########################

# Preparing counters for statistics 
num_male = 0                               # total male/female, including cocm staff
num_female = 0

num_male_member = [0]*4   # male/female member for each occupation status 
num_female_member = [0]*4
num_occ = [0]*4

num_havebeentococm = 0
num_cloth = [0]*6

num_email = 0
num_request = 0
num_allergy = 0
num_scholar = 0

num_serve = [0]*len(serve_sheet)

group=[[] for _ in range(num_group)]
Gleader_index = []

yr5 = []
yr4 = []
yr2 = []
yr1 = []
yr05 = []
yr0 = []
cocm = []

time_believe_group = [yr5,yr4,yr2,yr1,yr05,yr0]

###############################################################################

############################## Main code/body #################################
#--------------- This part is for generating two excel files -----------------#

# Header & styles 

for i in range(len(request)):
    Request.cell(row = 1, column=i+1).value= request[i]
    Request.cell(row = 1, column=i+1).fill = PatternFill(start_color = '96CDCD',fill_type='solid')
for i in range(len(allergy)):
    Allergy.cell(row = 1, column=i+1).value= allergy[i]
    Allergy.cell(row = 1, column=i+1).fill = PatternFill(start_color = '96CDCD',fill_type='solid')
for i in range(len(scholar)):
    Scholar.cell(row = 1, column=i+1).value= scholar[i]
    Scholar.cell(row = 1, column=i+1).fill = PatternFill(start_color = '96CDCD',fill_type='solid')
for i in range(len(email)):
    Email.cell(row = 1, column=i+2).value= email[i]
    Email.cell(row = 1, column=i+2).fill = PatternFill(start_color = '96CDCD',fill_type='solid')
camp_title = '营会'
Email.cell(1,1).value = camp_title

for i in range(len(serve_sheet)):
    #print(serve_sheet[i])
    for j in range(len(columns)):
        serve_sheet[i].cell(row = 1, column=j+1).value= columns[j]
        #print( serve_sheet[i].cell(row = 1, column=j+1).value)
        serve_sheet[i].cell(row = 1, column=j+1).fill = PatternFill(start_color = '96CDCD',fill_type='solid')

Group.cell(row=1, column=1).value = '组名'
Group.cell(row=1, column=2).value = '小组长1'
Group.cell(row=1, column=3).value = '小组长2'
Group.cell(row=1, column=2).font = Font(size=18,bold=True)
Group.cell(row=1, column=3).font = Font(size=18,bold=True)
Group.cell(row=1, column=2).alignment = Alignment(shrink_to_fit=True)
Group.cell(row=1, column=3).alignment = Alignment(shrink_to_fit=True)
Group.cell(row=1 + 1 + num_group, column=1).value = '不分组'
Group.cell(row=1 + 1 + num_group, column=1).fill = PatternFill(start_color = '880808',fill_type='solid')

# Save the excel files-----------------------------------------------------------------------------------------
dff.save(newbook)    # general classification
sdff.save(serbook)   # sevring team

#----------------------------> Anlysing the Name raw data
for i in range(num_row):
    for j in columns:
#------------------> This part is for general classification 
# Some counters for statistics
        # Counters 
        if j == '性别':                        # counter for males and females
            if df[j][i] == '女':
                num_female += 1
                if df['职业状态'][i] == '学生':
                    num_occ[0] += 1
                    num_female_member[0] += 1
                elif df['职业状态'][i] == '在职':
                    num_occ[1] += 1
                    num_female_member[1] += 1
                elif df['职业状态'][i] == 'COCM同工':
                    num_occ[2] += 1
                    num_female_member[2] += 1
                else:
                    num_occ[3] += 1
                    num_female_member[3] += 1
            else:
                num_male += 1
                
                if df['职业状态'][i] == '学生':
                    num_occ[0] += 1
                    num_male_member[0] += 1
                elif df['职业状态'][i] == '在职':
                    num_occ[1] += 1
                    num_male_member[1] += 1
                elif df['职业状态'][i] == 'COCM同工':
                    num_occ[2] += 1
                    num_male_member[2] += 1
                else:
                    num_occ[3] += 1
                    num_male_member[3] += 1

        elif j == '是否曾参加过COCM的大型营会':   # counter for have been to cocm
            if df[j][i] == '是，参加过':
                num_havebeentococm += 1
       
        elif j == '衣服尺寸':                    # counter for cloth sizes
            for k in range(len(size)):
                if df[j][i] == size[k]:
                    num_cloth[k] += 1
                    
# Re-arrange data into different sheets             
        # Producing new sheets & assosiate counters (easy ones)
        elif j == '未来是否愿意通过此邮箱继续收到COCM事工更新信息':
            
            if df[j][i] == '是':
                col = 0
                for k in email:
                    Email.cell(row = num_email+2, column=1).value = camp_name
                    Email.cell(row = num_email+2, column=col+2).value= df[k][i]
                    col += 1
                num_email += 1
        
        elif j == '备注':
            if type(df[j][i]) is str:
                col = 0
                for k in request :
                    Request.cell(row = num_request+2, column=col+1).value= df[k][i]
                    col += 1
                num_request += 1
                
        elif j == '食物过敏':
            if type(df[j][i]) is str:
                col = 0
                for k in allergy :
                    Allergy.cell(row = num_allergy+2, column=col+1).value= df[k][i]
                    col += 1
                num_allergy += 1
        
        elif j == '是否申请营费补助':
            if df[j][i] == '是':
                col = 0
                for k in scholar :
                    Scholar.cell(row = num_scholar+2, column=col+1).value= df[k][i]
                    col += 1
                num_scholar += 1
                
#----------------------> This part is for serving team
        elif j == '愿意在营会服事的岗位':
            if type(df[j][i]) != float :           # differetiate either serving or not
                ser_inter = df[j][i].split('\n') 
                
                if len(ser_inter) >= max_serve:    # highlight the potential overloads
                    #print(ser_inter)
                    for m in ser_inter:
                        num_False = 0                  # counter for anomaly, local counter
                        
                        for k in range(len(test_serve)):       
                            #print('m=',m,'k=',test_serve[k], m == test_serve[k])
                            if m == test_serve[k]:
                                if m == test_serve[0]:
                                    Gleader_index.append(i)
                                for n in range(len(columns)):
                                    #print(n,df['中文姓名'][i])
                                    #print(num_serve)
                                    serve_sheet[k].cell(row= num_serve[k]+2, column=n+1).value = df[columns[n]][i]
                                    serve_sheet[k].cell(row= num_serve[k]+2, column=n+1).fill = PatternFill(start_color = 'FFFF00',fill_type='solid')
                                num_serve[k] +=1
                            else:
                                num_False += 1
                                if num_False == len(test_serve):
                                    #print(df['中文姓名'][i])
                                    #print('Revisit',len(ser_inter),num_False)
                                    for n in range(len(columns)):
                                        serve_sheet[-1].cell(row= num_serve[-1]+2, column=n+1).value = df[columns[n]][i]
                                        serve_sheet[-1].cell(row= num_serve[-1]+2, column=n+1).fill = PatternFill(start_color = 'FFFF00',fill_type='solid')
                                    num_serve[-1] +=1
                else:
                    #print(ser_inter)
                    for m in ser_inter:
                        num_False = 0                  # counter for anomaly 
                        
                        for k in range(len(test_serve)):       
                            #print('m=',m,'k=',test_serve[k], m == test_serve[k])
                            if m == test_serve[k]:
                                if m == test_serve[0]:
                                    Gleader_index.append(i)
                                for n in range(len(columns)):
                                    #print(n,df['中文姓名'][i])
                                    #print(num_serve)
                                    serve_sheet[k].cell(row= num_serve[k]+2, column=n+1).value = df[columns[n]][i]
                                num_serve[k] +=1
                            else:
                                num_False += 1
                                if num_False == len(test_serve):
                                    #print(df['中文姓名'][i])
                                    #print('Revisit',len(ser_inter),num_False)
                                    for n in range(len(columns)):
                                        serve_sheet[-1].cell(row= num_serve[-1]+2, column=n+1).value = df[columns[n]][i]
                                    num_serve[-1] +=1
                                    
#--------------------> Group and Room allocators (harder ones)

# Prepare Group member indices by dublicating the overall indices
Gmember_index = []
for i in range(num_row):
    Gmember_index.append(i)

# Pick out the group members from the data, taking the Gleader indices away
index_correct = 0
for i in Gleader_index:
    Gmember_index.pop(i - index_correct )
    index_correct += 1

# Store group members into different year group
for i in Gmember_index:
    if df['信主时间'][i] == time_believe[0] :
        if df['性别'][i] == '男':
            yr5.insert(0,i)
        elif df['性别'][i] == '女':
            yr5.append(i)
    elif df['信主时间'][i] == time_believe[1] :
        if df['性别'][i] == '男':
            yr4.insert(0,i)
        elif df['性别'][i] == '女':
            yr4.append(i)
    elif df['信主时间'][i] == time_believe[2] :
        if df['性别'][i] == '男':
            yr2.insert(0,i)
        elif df['性别'][i] == '女':
            yr2.append(i)
    elif df['信主时间'][i] == time_believe[3] :
        if df['性别'][i] == '男':
            yr1.insert(0,i)
        elif df['性别'][i] == '女':
            yr1.append(i)
    elif df['信主时间'][i] == time_believe[4] :
        if df['性别'][i] == '男':
            yr05.insert(0,i)
        elif df['性别'][i] == '女':
            yr05.append(i)
    elif df['信主时间'][i] == time_believe[5] :
        if df['性别'][i] == '男':
            yr0.insert(0,i)
        elif df['性别'][i] == '女':
            yr0.append(i)
    else:
        cocm.append(i)

# Store the time believe groups data
time_believe_group_backup = time_believe_group.copy() 

# Well mix the data 
Gmember_index=[]
for i in range(len(time_believe_group)):           
   # random.shuffle(time_believe_group[i])             # randomise the data
   # random.shuffle(time_believe_group[i]) 
   # random.shuffle(time_believe_group[i]) 
    Gmember_index += time_believe_group[i]            # create new Gmember_index with random order

# Adding the well mixed members to each group in the order of the have-believed time, from longest to shortest
index_roll = 0
num_group_member = len(Gmember_index)
while index_roll < num_group_member:
    for i in range(num_group):
        if index_roll == num_group_member :
            break
        else:
            group[i].append(Gmember_index[index_roll])
            index_roll += 1
       
# test   
#a = 0
#for i in range(len(group)):
#    print(len(group[i]))
#    a += len(group[i])     

# Allocating the Group members
group_row = 0
for i in range(len(group)):
    for j in range(len(group[i])):
        index = group[i][j]
        Group.cell(row= group_row+2, column= j+4).value = ('%s %s %s %s %s'%(\
                                       df['ID'][index],df['中文姓名'][index],\
                        gender_trans(df['性别'][index]),df['信主时间'][index],\
                                       person_trans(df['性格倾向'][index])) )
        Group.cell(row= group_row+2, column= j+4).alignment = Alignment(vertical='center', wrap_text=True)
        if df['性别'][index] == '男':
            Group.cell(row= group_row+2, column= j+4).fill = PatternFill(start_color = '89CFF0',fill_type='solid')
        else :
            Group.cell(row= group_row+2, column= j+4).fill = PatternFill(start_color = 'ffc0cb',fill_type='solid')
    group_row += 1

# Insert cocm members alongside
cocm_row = 1
cocm_col = len(group[0])+14
for i in range(len(cocm)):
    index = cocm[i]
    cell = Group.cell(row=cocm_row ,column= cocm_col+i)
    
    if df['性别'][index] == '男' or df['性别'][index] == '女':   

        if ((len(group[0])+9+i) - cocm_col )%5 == 0:
            cocm_row += 1 
            cocm_col -= 5
            #print(i,cocm_row)
        cell.value = ('%s %s %s %s %s'%(\
                                   df['ID'][index],df['中文姓名'][index],\
                    gender_trans(df['性别'][index]),df['信主时间'][index],\
                                   person_trans(df['性格倾向'][index])) )
        
        cell.alignment = Alignment(vertical='center', wrap_text=True)
        if df['性别'][index] == '男':
            cell.fill = PatternFill(start_color = '89CFF0',fill_type='solid')
        elif df['性别'][index] == '女' :
            cell.fill = PatternFill(start_color = 'ffc0cb',fill_type='solid')
        else:
            None

# Save the excel files--------------------------------------------------------------------------
dff.save(newbook)    # general classification
sdff.save(serbook)   # sevring team
#-----------------------------------------------------------------------------#

# Statistics output in txt
f = open(newtxt,'w',encoding = "utf-8") #encoding is for chinese character

total_num = num_male + num_female

f.write("总人数（包括cocm同工） : %d"%(total_num))
f.write("\n男生人数 （包括cocm同工）: %d; 男学生: %d; 男在职: %d; 男中心同工: %d; 男其他: %d" %(num_male,num_male_member[0],num_male_member[1],\
                                                                        num_male_member[2],num_male_member[3]))
f.write("\n女生人数 （包括cocm同工）: %d; 女学生: %d; 女在职: %d; 女中心同工: %d; 女其他: %d"%(num_female,num_female_member[0],num_female_member[1],\
                                                                        num_female_member[2],num_female_member[3]))
f.write("\n经费资助 : %d"%(num_scholar))
f.write("\n之前参加过营会的人 : %d"%(num_havebeentococm))
f.write("\n\n\n衣服 :")
f.write("\n%s : %d ;\n%s : %d ;\n%s : %d ;\n%s : %d ;\n%s : %d ;\n%s : %d ;"%(size[0],num_cloth[0],\
                                    size[1],num_cloth[1],size[2],num_cloth[2],size[3],num_cloth[3],\
                                    size[4],num_cloth[4],size[5],num_cloth[5]))
f.write("\n\n\n服侍情况 ：")
f.write("\n检查Revisit sheet, 有人可能填的东西有问题 或者 sheet title name 需要更换")
f.write("\n检查sheet header section，核对一下服侍的heaers")
f.write("\nSheets 中用黄色标出来的是有服侍超过%d个的\n\n\n"%(max_serve))
f.write("服侍统计 ：")
for i in range(len(test_serve)):
    f.write("\n%s : %d"%(test_serve[i],num_serve[i]))
f.close()

