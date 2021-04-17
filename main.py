import pandas as pd
import datetime
import smtplib
import json

# # Enter Your authantication details in Config.json
with open('config.json', 'r') as c:
    params = json.load(c)["params"]


def SendEmail(to, Cname, admiNo, gmail, Ymsg, msg):
    print(f"{to}, {Cname}, {admiNo}, {gmail}, {Ymsg}, {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(params['GMAIL_ID'], params['GMAIL_PASSWORD'])
    s.sendmail(params['GMAIL_ID'], gmail, f"Subject: {Ymsg}\n\n {to}\n{Cname}\n{admiNo}\n{msg}")
    s.quit()

if __name__ == '__main__':
    df = pd.read_excel("maildata.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m-%Y")
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(type(today))
    writeInd = []
    for index, item in df.iterrows():
        # print(index, item['Name'])
        if yearNow not in str(item['Send']):
            SendEmail(item['Name'], item['CName'], item['AdmissionNumber'],
                      item['GmailId'], item['ExtraMessage'], item['Message'])
            writeInd.append(index)

    if writeInd == []:
        exit()

    for i in writeInd:
        yr = df.loc[i, 'Send']
        # print(yr)
        df.loc[i, 'Send'] = str(yr) + ', ' + str(yearNow)
        # print(df.loc[i, 'Send'])
    df.to_excel('maildata.xlsx', index=False)











