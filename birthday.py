import pandas as pd
import datetime
import smtplib
import os
os.chdir(r"E:\birthday_wisher")
os.mkdir("testing")
#Enter your Authentication details
GMAIL_ID='lofoproject10@gmail.com'
GMAIL_PSWD='vyqjxebvraklervt'
def sendMail(to,sub,msg):
    print(f"Email to {to} sent with subj {sub} with msg {msg}")
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PSWD)
    s.sendmail(GMAIL_ID,to,f"Subject:{sub}\n\n\n{msg}")
    s.quit()

if __name__ == "__main__":
    # sendMail('deepapandey364@gmail.com',"subject","TEst message")
    # exit()
    df=pd.read_excel("data.xlsx")
    #print(df)
    today=datetime.datetime.now().strftime("%d-%m")
    yrNow=datetime.datetime.now().strftime("%Y")
    #print(type(today))
    writeInd=[]
    for index,items in df.iterrows():
        bday=items['BIRTHDAY'].strftime("%d-%m")
        #yrNow=items['BIRTHDAY'].strftime("%y")
        #print(bday)
        if(today==bday) and yrNow not in str(items['YEAR']):
            #sendMail(items['Email'],"Happy b'day - Deepa",items['DIALOUGE']) 
            #sendMail('deepapandey364@gmail.com',"Happy b'day - Deepa",items['DIALOUGE']) 
            # print(items['Email'])
            # print(items['YEAR'])
            writeInd.append(index)
    print(writeInd)
        
    for i in writeInd:
        yr=df.loc[i,'YEAR']
        df.loc[i,'YEAR']=str(yr)+' , '+str(yrNow)
        print(df.loc[i,'YEAR'])
    #print(df)
        df.to_excel('data.xlsx',index=False)