import pandas as pd # to read csv and excel files
import datetime # to deal with date and time
import smtplib # used to send emails
import os
os.chdir(r"C:\Users\Mani\birthday_wish")
#os.mkdir("testing")

# Enter your authentication details

GMAIL_ID = '' # Add your mail id
GMAIL_PSWD = '' # Add your password
def sendEmail(to,sub,msg):
    print(f"Email to {to} sent with subbject:{sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to ,f"Subject:{sub}\n\n{msg}") # from,to,sub,msg
    s.quit()

if __name__ == "__main__":
    #sendEmail(GMAIL_ID,"subject","test message")
    #exit()
    df=pd.read_excel("data.xlsx") # sample values are taken take values on your own
    #print(df)
    today = datetime.datetime.now().strftime("%d- %m")# today in string format
    yearNow = datetime.datetime.now().strftime("%Y")
    #print(today)
    writeInd= []

    for index,item in df.iterrows():
        #print(index,item['Birthday'])
        bday = item['Birthday'].strftime("%d- %m")
        #print(bday)
        if(today == bday) and yearNow not in str(item["Year"]):
            sendEmail(item['Email'],"Happy Birthday",item['Dialogue'])
            writeInd.append(index) # index where I have sent the mail
    print(writeInd)

    if writeInd != []:
        for i in writeInd:
            yr = df.loc[i,'Year'] # uisng row and column index
            #print(yr)
            df.loc[i,'Year'] = str(yr) + ',' + str(yearNow) # like 2019 ,2020
            #print(df.loc[i,'Year'] )

        #print(df)
        df.to_excel('data.xlsx',index=False)# don't want to write index