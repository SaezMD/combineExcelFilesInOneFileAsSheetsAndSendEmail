# Add multiple excel files as sheets in one workbook and send email usign Outlook
import pandas as pd
import glob, os, datetime, time
import win32com.client as win32
#from datetime import datetime

print('Starting Combine Microsoft and Oracle Files')

start_time = time.time()
# Today's date to open the correct file
todayDate = ' ' + datetime.date.today().strftime("%d%m%Y")
print('[MICRO] - Today is:' + todayDate)

dest_dir = os.path.normpath(os.path.expanduser("~/Desktop"))

# Files to combine (2) based on today's date in files names. Can add more
excel_files = glob.glob(dest_dir +'/*' + todayDate + ' PYTHONtoSendTo_FTP.xlsx') + glob.glob(dest_dir + '/INFO MICRO' + todayDate + '*')
print(f'[MICRO] - Files to merge: {excel_files} ')

#Create destination file:
destination = dest_dir + '\Stock MICRO' + todayDate + '.xlsx'
print(f'[MICRO] - Destination: {destination} ') # Destination INFO
writer = pd.ExcelWriter(destination,engine='xlsxwriter')

resultSheets = ["Microsoft","Oracle"] #Sheets names after joint the files

#Combinator of files in sheet:
print(f'[MICRO] - Combining sheets... ')
x = 0
for excel_file in excel_files:
    sheet = resultSheets[x]
    print(f' Sheet {x}: {sheet} ')
    df1 = pd.read_excel(excel_file)
    df1.fillna(value='N/A', inplace=True)
    df1.to_excel(writer, sheet_name=sheet, index=False)
    #x = x + 1 long form
    x += 1

writer.save()
print(df1) # Check the data

# ------------------ Send email code part

# Check if Outlook is ready:
def outlook_is_running():
    import win32ui
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False

if not outlook_is_running():
    import os
    os.startfile("outlook")

# reading the spreadsheet for emails
email_list = pd.read_excel('D:/Python/New FILES/Filters/Emails.xlsx')
  
# getting the names and the emails
emails = email_list['Email']
emailforOutlook = "" 
# iterate through the records
for i in range(len(emails)):
    # for every record get the email addresses
    email = emails[i]
    emailforOutlook = emailforOutlook + ";" + email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.Subject = 'STOCK MICRO' + todayDate
mail.BCC = emailforOutlook
mail.HTMLBody = r"""
Dear all,<br><br>
Please, find attached today's report.<br><br>
In case you need further details, do not hesitate to contact me: saulsaezrodriguez@gmail.com<br><br>
Best regards,<br>
"""
# Attach files to the email
mail.Attachments.Add(os.path.normpath(destination))

print(f'[MICRO] - Email ready!')
mail.Send()
print(f'[MICRO] - Email sent!')

print(f'[MICRO] - Done! Completed in {round(time.time()-start_time,2)} seconds.')