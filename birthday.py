import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"E:\birthday_wisher") #Schedule the task scheduler to wish on the specified date & give proper file location 
#os.mkdir("testing")
#Enter your Authentication details
GMAIL_ID=''
GMAIL_PSWD=''
def sendMail(to,sub,msg):
    print(f"Email to {to} sent with subj {sub} with msg {msg}")
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to,f"Subject:{sub}\n\n\n{msg}")
    s.quit()

if __name__ == "__main__":
    #perform pip install xlrd
    df=pd.read_excel("data.xlsx")
    today=datetime.datetime.now().strftime("%d-%m")
    yrNow=datetime.datetime.now().strftime("%Y")
    writeInd=[]
    for index,items in df.iterrows():
        bday=items['BIRTHDAY'].strftime("%d-%m")
        if(today==bday) and yrNow not in str(items['YEAR']):
            #sendMail(items['Email'],"Happy b'day - Deepa",items['DIALOUGE'])
            writeInd.append(index)
    #To keep a track of already wished b'days on that particular year
    for i in writeInd:
        yr=df.loc[i,'YEAR']
        df.loc[i,'YEAR']=str(yr)+' , '+str(yrNow)
        print(df.loc[i,'YEAR'])
        #perform pip install openpyxl before it
        df.to_excel('data.xlsx',index=False)
