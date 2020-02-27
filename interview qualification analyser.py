import xlrd
import re 


#function cheacking substring present or not

def check(string, sub_str): 
    if (string.find(sub_str) == -1): 
        return("NO") 
    else: 
        return("YES") 
            

#reading excel data 

loc = ("Data_Science_2020_v2.xlsx") 

wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)


#function it giving score based on bachelors completed year
def bachelor_rating(x):
    if x=="2020" or "2021":
        return 10
    elif x=="2019":
        return 8
    else:
        return 5

#function it giving score based on masters completed year

def masters_rating(x):
    if x=="2020" or "2021":
        return 7
    else:
        return 3


    
# cheacking applicant maximum qualification is bachelors or not

def accademic_bachelor(x):
    flag=0
    y=sheet.cell_value(x, 7)
    bachelors=["B.E","B.Tech"]
    for i in bachelors:
        if check(y, i)=="YES":
            flag=flag+1
    if flag>0:
        return bachelor_rating(x)
    else:
        return 0


# cheacking applicant maximum qualification is masters or not 

def accademic_masters(x):
    flag=0
    y=sheet.cell_value(x, 7)
    bachelors=["MSc","M.Tech"]
    for i in bachelors:
        if check(y, i)=="YES":
            flag=flag+1
    if flag>0:
       return masters_rating(x)
    else:
        return 0


#switch function it gives score self rating for each language
def self_rating_score(i):
        switcher={
                1:3,
                2:7,
                3:10,
                
             }
        return( switcher.get(i,0))




# function it call score for self rated 3 subjects
def self_rating(x):
    rate=0
    for j in range(2,5):
        rate=rate+self_rating_score(sheet.cell_value(x, j))
    return rate




# function gives 3 marks for required skills
def skill(x):
    # input comma separated values as string 
    data=sheet.cell_value(x, 5)
    score=0
    # convert to the list
    required_skill_list=["Machine Learning", "Deep Learning", "Natural Language Processing(NLP)", "Statistical Data Analysis", "Statistical Modeling", "SQL","NoSQL", "Amazon Web Services(AWS)","MS-Excel"]
    
    for i in required_skill_list:
    
        if (data.find(i)>=0): 
            score=score+3
        else: 
            score=score+0
    return score
            



# calling each row and evaluating score for each row
qualified_list=[]
shr = wb.sheet_by_name('Sheet1')
for rownum in range(1,shr.nrows):
    score=0
    score=score+accademic_bachelor(rownum)
    score=score+accademic_masters(rownum)
    score=score+self_rating(rownum)
    score=score+skill(rownum)
    if score>=45:
        qualified_list.append((sheet.cell_value(rownum, 0),score))



#printing names and score of qualified applicants

print("Qualified applicant and their score are:")    
for i in qualified_list:
    print (i)





    
    
        
#print(sheet.cell_value(1, 7))

#accademic_bachelor(str(sheet.cell_value(1, 7)))
