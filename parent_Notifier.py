# importing the nessesary packages
import schedule
import time
import smtplib
import pandas as pd
import gspread


def Main():
    # Authenticating Sheets
    # Calling the Service_Account File
    sa = gspread.service_account(filename="service_account.json")
    sh = sa.open("Attendence")  # Name of the Google Sheet
    wks = sh.worksheet("Master Data")  # Name of the Master Data WorkBook
    wks2 = sh.worksheet("Daily Entry")  # Name of the Daily Entry WorkBook

    # Importing and Cleaning The Data
    df_master = pd.DataFrame(wks.get_all_records(), index=None)
    df_daily_dirty = pd.DataFrame(wks2.get_all_records(), index=None)
    df_daily = df_daily_dirty.drop_duplicates(subset=["Barcode"])

    # Checking For Data Availability
    if(df_daily.empty == False):
        df_data = df_master.merge(df_daily, on='Barcode')

        #>>> Emailing Department <<<#
        # Credentials
        your_email = "YOUR_EMAIL@GMAIL.COM"  # The sender's Email
        your_password = "YOUR_APP_PASSWORD"  # The App Password

        # Establishing connection with gmail
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(your_email, your_password)

        # Sending Emails
        for row in range(df_data.count(0).Barcode):
            barcode = df_data.Barcode[row]
            name = df_data.Name[row]
            parentname = df_data.Parent[row]
            email = df_data.Email[row]
            mailsent = df_data.Mailsent[row]
            subject = f'{name} Have Reached School (IISJED)'
            body = f'Hello {parentname}, {name} Has Reached School'
            message = f'Subject: {subject},\n\n{body}'

            # Check If Email Already Sent
            if(mailsent == ""):
                cell = wks2.find(str(barcode))
                wks2.update_cell(cell.row, 2, "Y")
                print(f'Email Sent to {parentname} Saying ( {message} )')
                server.sendmail(your_email, [email], message)
            else:
                print("All The Emails Are Sent!")
        server.close()


print("Starting The Script")
schedule.every(5).seconds.do(Main)

while True:
    schedule.run_pending()
    time.sleep(1)
