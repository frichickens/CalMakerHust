import pandas as pd

timeTable = pd.read_excel(r"timeTable.xlsx")
registedClass = pd.read_excel(r"registedClass.xlsx")
week = timeTable['Week']
day  = timeTable['Day']
time  = timeTable['Time']
location = timeTable['Location']
className = timeTable['Class Name']
code =  registedClass['Code']
subCode = registedClass['Sub Code']
name = registedClass['Name']
classType = registedClass['Type']

newWeek = []
f = open("FinalRes.txt", "w",encoding="utf8")

def checkComma(arr):
    for x in arr:
        if(x=='-'):
            return True
            break
        else:
            pass
    return False

def findPos(a,code,subCode):
    for x in range(len(code)):
        if(code[x]==a):
            return x
    for x in range(len(code)):
        if(subCode[x]==a):
            return x

maxWeek = 0
minWeek = 100


for x in range(len(week)):
    arr = []
    if((len(str(week[x])))>2):
        tmp = ''
        if(checkComma(str(week[x]))):
            arr = str(week[x]).replace(",", "-")
            arr = arr.split("-")
            if(len(arr)%2==0):
                for y in range(0,len(arr),2):
                    tmp += " ".join(str(i) for i in range(int(arr[y]),int(arr[y+1])+1)) +" "
            else:
                for y in range(0,len(arr)-1,2):
                    tmp += " ".join(str(i) for i in range(int(arr[y]),int(arr[y+1])+1)) +" "
                tmp += str(arr[-1])
        else:
            arr = str(week[x]).split(",")
            tmp  = " ".join(arr)
        temp = int(arr[-1])
        temp1 = int(arr[0])
        if(temp>maxWeek): maxWeek = temp
        if(minWeek>temp1): minWeek = temp1
        newWeek.append(tmp)
    else:
        newWeek.append(str(week[x]))
        if(week[x]>maxWeek):
            maxWeek = week[x]
        if(minWeek>week[x]):
            minWeek = week[x]

weekLength = [int(x) for x in range(minWeek,maxWeek+1)]
cnt = 0
pos = 0
finalRes = []
num = []
while (cnt!=len(weekLength)):
    txt = []
    f.write('Week '+str(weekLength[cnt]) + ': \n')
    for x in range(len(week)):
        num = [int(s) for s in newWeek[x].split()]
        if(weekLength[cnt] in num):
            pos = findPos(className[x],code,subCode)
            f.write(str(day[x])+ ' | ' + str(location[x]) + ' | '+str(time[x]) + ' | ' + 
                    str(className[x]) + ' | '+ str(name[pos]) + ' | '+ str(classType[pos])  + '\n')
    cnt+=1
f.close()
