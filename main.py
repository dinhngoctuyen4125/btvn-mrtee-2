import openpyxl

book = openpyxl.load_workbook('Ex2.xlsx')

content_sheet = book['Content']
mail_sheet = book['Mail']

contents = []
mails = []

inx = 2
while(True):
    val = content_sheet['B' + str(inx)].value
    if(val == None): break
    contents.append(val)
    inx+=1

inx = 2
while(True):
    val = mail_sheet['B' + str(inx)].value
    if(val == None): break
    mails.append(val)
    inx+=1

book = openpyxl.Workbook()

sheet = book.active
sheet.title = 'Merge'
sheet.append(['STT', 'Nội dung', 'Mail'])

stt = 1
for content in contents:
    for mail in mails:
        sheet.append([stt, content, mail])
        stt+=1

book.save('new_Ex2.xlsx')
