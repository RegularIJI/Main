import openpyxl
import os
import csv

with open('Приведи друга в УП.csv','w',newline ='') as file:
    writer = csv.writer(file, delimiter =";")



fileslist = []

Final_tsk = [['surname ','name ','patronymic ','phone_mobile ','email','Who_recommended','Логин пригласившего','vacancy','City','Functional']]

Family = []
Name = []
Otchestvo = []
Phone = []
Mail = []
City = []
Employee_FIO = []
Employee_ID = []
Functional = []
Vacancy = []
Crack = []
Crack_score = 0
a = os.getcwd()
files = os.listdir(a)
for i in range(0,len(files)):
    if len(files[i]) >= 4 and (files[i][-4:]) == 'xlsx' and files[i] != 'Ass.xlsx':
        fileslist.append(files[i])
for i in range(0,len(fileslist)):
    book = openpyxl.open(fileslist[i],read_only =True)
    sheet = book.active
    Family.append(sheet[3][2].value)
    Name.append(sheet[4][2].value)
    Otchestvo.append(sheet[5][2].value)
    Phone.append(sheet[6][2].value)
    Mail.append(sheet[7][2].value)
    City.append(sheet[8][2].value)
    Employee_FIO.append(sheet[9][2].value)
    Employee_ID.append(sheet[10][2].value)
    Functional.append(sheet[11][2].value)
    Vacancy.append(sheet[12][2].value)

count = len(Family)
print (count)
i = count -1

j = i

while i != -1:
    if type(Phone[i]) != str:
        del (Family[i])
        del (Name[i])
        del (Otchestvo[i])
        del (Phone[i])
        del (Mail[i])
        del (City[i])
        del (Employee_FIO[i])
        del (Employee_ID[i])
        del (Functional[i])
        del (Vacancy[i])
        Crack_score = 1
        Crack.append(fileslist[i])
        i = i - 1
    else:
        b = (Phone[i]).split()
        b = ''.join(b)
        Phone[i] = b
        if (len(Phone[i]) != 12) or (Phone[i] == '+79991234567'):
            del (Family[i])
            del (Name[i])
            del (Otchestvo[i])
            del (Phone[i])
            del (Mail[i])
            del (City[i])
            del (Employee_FIO[i])
            del (Employee_ID[i])
            del (Functional[i])
            del (Vacancy[i])
            Crack_score = 1
            Crack.append(fileslist[i])
            i = i - 1
        else:
            i = i - 1






count = len(Family)
for i in range(0,count):
  f = []
  f.append(Family[i])
  f.append(Name[i])
  f.append(Otchestvo[i])
  f.append(Phone[i])
  f.append(Mail[i])
  f.append(Employee_FIO[i])
  f.append(Employee_ID[i])
  f.append(Vacancy[i])
  f.append(City[i])
  f.append(Functional[i])
  Final_tsk.append(f)

if Crack_score == 1:
    print('Не обработано: ',Crack)
    a = input('Для завершения программы введите любой символ')


with open('Приведи друга в УП.csv','w',newline ='') as file:
    writer = csv.writer(file, delimiter =";")
    writer.writerows(Final_tsk)
