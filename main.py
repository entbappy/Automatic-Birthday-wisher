"""
Project: Automatic Birthday Wisher 
Author : Bappy Ahmed 
Email : bappymalik4161@gmail.com
Date: 03 oct 2020
"""
import pandas as pd 
import datetime
import smtplib

# Enter your gmail and password
GMAIL_ID = ''
GMAIL_PASS = ''

#This function sends mail/sms
def sendmail(to,sub,msg):
    print(f"Email to {to} sent with subject {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PASS)
    s.sendmail(GMAIL_ID,to, f"Subject: {sub}\n\n{msg}")
    s.quit()

if __name__ == "__main__":
    #Testing purpose
    # to = 'bappymalik4161@gmail.com'
    # sendmail(to,"Hey","Whats up!")
    # exit()
    
    df = pd.read_excel('data.xlsx')
    # print(df)
    today = datetime.datetime.now().strftime('%d-%m')
    yearnow = datetime.datetime.now().strftime('%Y')
    # print(yearnow)

    update =[]

    #Taking data form the dataset
    for index,item in df.iterrows():
        # print(index)
        # print(index,item['Birthday'])
        birth = item['Birthday'].strftime('%d-%m')
        # print(birth)
        
        
        if (birth == today) and yearnow not in str(item['Year']):
            sendmail(item['Email'],'Happy Birthday!',item['Dialouge'])
            update.append(index)
    print(update)


    for i in update:
        year = df.loc[i,'Year']
        df.loc[i,'Year'] = str(year) + ',' + str(yearnow)
        # print(df.loc[i,'Year'] )
    # print(df)

    df.to_excel('data.xlsx',index=False)

 

        
