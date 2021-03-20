import pandas as pd
from openpyxl import load_workbook
z = pd.read_excel('pythonexcel.xlsx', sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'])
tmp = []
n = int(input('Enter no'))
for _ in range(n):
    tmp1 = []
    h = int(input('Enter your ps no:'))
    name = str(input('Enter the Name:'))
    email = str(input('Enter the email'))
    tmp1.append(h)
    tmp1.append(name)
    tmp1.append(email)
    tmp.append(tmp1)
df1 = pd.DataFrame(columns=['SL#', 'PS number', 'Display Name', 'Official Email Address',
       'company name', 'Year of join', 'Room No', 'Block', 'Area', 'location',
       'Training Room', 'Training Name', 'Domain', 'c lang', 'linux', 'python',
       'Salery', 'Designation', 'Blood group', 'Gender', 'Phone no',
       'Adhar Num', 'In Time', 'Out Time', 'BUS NUM', 'Attendence',
       'Self decleration', 'Temperature', 'Age', 'Marital status', 'DOB',
       'Country', 'State', 'Initial'])
for i in tmp:
    h, name, email = i
    y = z['Sheet1']
    y = y[(y['PS number'] == h) & (y['Display Name'] == name) & (y['Official Email Address'] == email)]
    if len(y) == 0:
        print('No match')
    else:
        df = pd.DataFrame(y, columns = ['SL#','PS number','Display Name','Official Email Address'])
        for i in z.keys():
            x = z[i]
            t = x[(x['PS number'] == h) & (x['Display Name'] == name) & (x['Official Email Address'] == email)]
            col = x.columns
            for j in col:
                df[j] = t[j]
    df1 = df1.append(df)

df2 = df1.describe()
book = load_workbook('pythonexcel.xlsx')
with pd.ExcelWriter('pythonexcel.xlsx', engine='openpyxl') as w:
    w.book = book
    w.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df1.to_excel(w, sheet_name='master', index=False)
    df2.to_excel(w, sheet_name='summary')
    w.save()
    w.close()
