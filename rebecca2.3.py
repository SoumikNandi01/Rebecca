#! python3
# Assistant

# Modules Needed

import pyttsx3
import sys,requests,shelve,pprint,webbrowser,time,pyperclip,openpyxl,imaplib,email,os,wikipedia,wikiquotes,speech_recognition as sr,pyaudio,random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import textblob
from textblob import TextBlob
from textblob.classifiers import NaiveBayesClassifier as nbc

engine = pyttsx3.init()
ids = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0"
rate = engine.getProperty("rate")
engine.setProperty("rate",rate-50)
engine.setProperty("voice",ids)
engine.say("Welcome User!. I am Rebecca. I can help you with your daily stuffs. HAPPY AUTOMATION")
engine.runAndWait()

#Test Corpus For NBC

happy_list = ["good to hear it","That's nice","I am happy that you are happy","that is great","I am so happy for you","always be happy like this","Good to hear that"]
sad_list = ["you look sad","things doesnt look good","everything gets okay with time","Cheer up","Please dont be sad","thats okay mate. life is so much more"]



test_corpus = [("Login facebook","login"),("login facebook","login"),("log account","login"),("login account","login"),("open facebook account","login"),("open account","login")
               ,("login account","login"),("login linkedin","login"),("login rediffmail","login"),("open linkedin","login"),("open rediffmail","login"),("open facebook","login")
               ,("log linkedin","login"),("open microsoft","login"),("login gmail","login"),("log gmail","login"),("log microsoft","login"),("open rediffmail","login")
               ,("open gmail","login"),("open linkedin","login"),("open gmx","login")
               ,("google search","search"),("search","search"),("look web","search"),("is meaning","search"),("is meant","search")
               ,("is","search"),("is the","search"),("Which is the","search"),("Where is","map"),("Where are","map"),("Tell who","search")
               ,("Tell me what","search"),("Can you tell me","search"),("Do you know","search"),("Do you now who","search"),("Can you find","search")
               ,("can you see","search"),("email to","email"),("send mail","email"),("drop mail","email"),("message","email"),("mail to","email")
               ,("emails do I","inbox"),("View emails","inbox"),("unread emails","inbox"),("Fetch unread emails","inbox"),("do i have inbox","inbox")
               ,("Check inbox mails","inbox"),("Check mails","inbox"),("New mails","inbox"),("check inbox","inbox"),("Show unread emails","inbox")
               ,("Show mailbox","inbox"),("update database","database"),("add database","database"),("add records","database"),("delete database","database")
               ,("Show database","database"),("show accounts","database"),("add accounts","database"),("delete accounts","database"),("are accounts","database")
               ,("delete accounts","database"),("wikipedia","wikipedia"),("search wiki","wikipedia"),("search wikipedia for","wikipedia"),("movie review","wikipedia"),("movie rating","wikipedia")
               ,("old is","wikipedia"),("who is","wikipedia"),("tell quote","quote"),("quote","quote"),("quotations","quote"),("quotation","quote"),("me","quote"),("tell","quote")
               ,("say quote","quote"),("say quotation","quote"),("quotations","quote"),("quote","quote"),("me quote","quote"),("maps","map"),("place","map"),("location","map")
               ,("located","map"),("locate","map"),("google maps","map"),("where is","map"),("where","map"),("where is","map"),("amazing","happy"),("won","happy"),("beautiful","happy")
               ,("amazing","happy"),("good","happy"),("great","happy"),("lovely","happy"),("feeling awesome","happy"),("pleasure","happy"),("bliss","happy"),("delight","happy"),("enjoy","happy")
               ,("joy","happy"),("cheerful","happy"),("laugh","happy"),("well-being","happy"),("prosperity","happy"),("cheer","happy"),("ecstacy","happy"),("rejoice","happy")
               ,("unhappy","sad"),("depression","sad"),("displeasure","sad"),("trouble","sad"),("worry","sad"),("upset","sad"),("sad","sad"),("misery","sad"),("pain","sad"),("sorry","sad")
               ,("trouble","sad"),("broke up","sad"),("sorry","sad"),("bad","sad"),("failed","sad"),("broke","sad"),("kicked","sad"),("not going well","sad")]

model = nbc(test_corpus)

#happy

def happy():
    x =random.randint(0,len(happy_list)-1)
    reply = happy_list[x]
    engine.say(reply)
    engine.runAndWait()

#sad

def sad():
    x = random.randint(0,len(sad_list)-1)
    reply = sad_list[x]
    engine.say(reply)
    engine.runAndWait()

#maps

def map_():
    engine.say("Which place do you wanna search")
    engine.runAndWait()
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source,duration=2)
        print("Speak Now")
        audio = r.listen(source)
    print("Recognising Audio")
    try:
        item = r.recognize_google(audio)
    except sr.UnknownValueError:
        engine.say("Sorry I did not get that")
        engine.runAndWait()
    browser = webdriver.Chrome()
    browser.get("https://google.com/maps/place/"+item)
    button = browser.find_element_by_xpath('//*[@id="searchbox-searchbutton"]')
    button.send_keys(Keys.ENTER)

    
#quote

def quote():
    engine.say("Here is a quote for you")
    engine.runAndWait()
    quote = wikiquotes.quote_of_the_day("english")
    quote = quote[0]
    engine.say(quote)
    engine.runAndWait()

#Wiki

def wiki():
    engine.say("What can i search for you in the wikipedia?")
    engine.runAndWait()
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source,duration=2)
        print("Speak Now")
        audio = r.listen(source)
    print("Recognising Audio")
    try:
        topic = r.recognize_google(audio)
    except sr.UnknownValueError:
        engine.say("Sorry! I did not get it")
        engine.runAndWait()
    text = wikipedia.summary(topic,sentences=3)
    engine.say("Here's what I got")
    engine.runAndWait()
    engine.say(text)
    engine.runAndWait()
    
# Adding Data To Database

def add_data():
    account = input("Account Name: ")
    username = input("Email/Username: ")
    passw = input("Passwords: ")
    doc = openpyxl.load_workbook("Database.xlsx")
    sheet = doc["database"]
    last_row = sheet.max_row
    sheet["B"+str(last_row+1)] = account.lower() #Column Number
    sheet["C"+str(last_row+1)] = username.lower() #Column Number
    sheet["D"+str(last_row+1)] = passw.lower() #Column Number
    doc.save("Database.xlsx") # Name of the database
    pprint.pprint("Database Uploaded")
    
# Deleting Data From Database
    
def delete():
    account = input("Account name to be deleted: ")
    doc = openpyxl.load_workbook("Database.xlsx")
    sheet = doc["database"]
    last_row = sheet.max_row 
    for i in range(3,last_row+1): #data starts from 3rd Row. You may change it later
        if sheet["B"+str(i)].value == account.lower(): #Column Number
            sheet.delete_rows(i,1)
            break
    doc.save("Database.xlsx")
    pprint.pprint("Database Uploaded")

# Viewing Data In The Database

def view():
    doc = openpyxl.load_workbook("Database.xlsx")
    sheet = doc["database"]
    for i in  range(3,sheet.max_row+1): #data starts from 3rd Row. You may change it later
        pprint.pprint(sheet["B"+str(i)].value) #Column Number
    doc.save("Database.xlsx")

# Sending Emails

def send_email():
    address = input("Recipent's Address : ")
    subject = input("Type subject of email: ")
    text = input("Type message body: ")
    import smtplib
    server = smtplib.SMTP_SSL("smtp server of your service provider",465)
    server.ehlo()
    server.login("your-username","your-password")
    server.sendmail("your-username",address,"Subject:"+subject+"\n"+text)
    engine.say("Done")
    engine.runAndWait()
    server.quit()

# Search Internet

def search_net():
    engine.say("What can I search for you?")
    engine.runAndWait()
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source,duration=2)
        print("Speak Now")
        audio = r.listen(source)
    print("Recognising Audio")
    try:
        item = r.recognize_google(audio)
        webbrowser.open("https://google.com/search?q="+item)
    except sr.UnknownValueError:
        engine.say("Sorry! I did not get that")
        engine.runAndWait()

# Receiving Mails
    
def mail():
    server = imaplib.IMAP4_SSL("imap server of your mail provider")
    server.login("your-username","your-password")
    server.select("INBOX")
    typ, data = server.search(None,"UNSEEN")
    keys = data[0].decode().split(" ")
    temp = keys[len(keys)-1]
    x = "s"
    if temp == '':
        engine.say("You have 0 unread emails")
        engine.runAndWait()
        server.logout()
        pprint.pprint("Logged out")
    else:
        engine.say("You have "+str(len(keys))+" unread emails")
        engine.runAndWait()
        i = int(temp)
        while int(i)>0:
            r = sr.Recognizer()
            mic = sr.Microphone()
            with mic as source:
                r.adjust_for_ambient_noise(source,duration=2)
                print("Speak Now")
                audio = r.listen(source)
            print("Recognising Audio")
            item = r.recognize_google(audio)
            item = item.lower()
            
            if "subject" in item:
                typ,data = server.fetch(str(i),"(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                subject = msg["subject"]
                engine.say(subject)
                engine.runAndWait()
            elif "from" in item:
                typ,data = server.fetch(str(i),"(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                frm = msg["from"]
                engine.say(frm)
                engine.runAndWait()
            elif "time" or "when" in item:
                typ,data = server.fetch(str(i),"(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                date = msg["date"]
                engine.say(date)
                engine.runAndWait()
            elif "next" in item:
                i = int(i)-1
                if(int(i)<=0):
                    engine.say("You do not have any unread emails now")
                    engine.runAndWait()
                    break
            elif "exit" or "logout" or "out" or "log" in item:
                break
            else:
                "INVALID COMMAND!"
        server.logout()
        engine.say("Logged out")
        engine.runAndWait()

# Getting FB notifications

def get_notifs_facebook():
    engine.say("Fetching notifications..Please wait")
    engine.runAndWait()
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    browser = webdriver.Chrome(options = chrome_options)
    browser.get("https://facebook.com/notifications")
    user = browser.find_element_by_id("email")
    user.send_keys("your-username")
    pwd = browser.find_element_by_id("pass")
    pwd.send_keys("your-password")
    pwd.send_keys(Keys.ENTER)
    for i in range(1,6):
        xpaths = '//*[@id="u_0_t"]/div/ul/li['+str(i)+']/div/div/a/div/div[2]/div/div/div[2]/div/div/span'
        try:
            notif = browser.find_element_by_xpath(xpaths)
        except:
            continue
        engine.say(notif.text)
        engine.runAndWait()
    

# Main Function Starts

while 1:
    engine.say("What can I do for you?")
    engine.runAndWait()
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source,duration=3)
        print("Speak Now")
        audio = r.listen(source)
    print("Recognising Audio")
    command = r.recognize_google(audio)
    command = command.lower()
    if "bye" in command:
        engine.say("Bye Have a nice day")
        engine.runAndWait()
        sys.exit()
    if "facebook" in command and ("notifications" in command or "notification" in command):
        get_notifs_facebook()
        continue
    if(command == "open the database file"):
        engine.say("Enter the password")
        engine.runAndWait()
        pwds = input()
        if(pwds == "your-password"):
            os.startfile("Database.xlsx")
            continue
        else:
            engine.say("Access Denied")
            engine.runAndWait()
            continue
    simplified = ""
    blob = TextBlob(command)
    for tag,pos in blob.tags:
        if "NN" in pos or "VB" in pos or "JJ" in pos:
            simplified=simplified+tag+" "
    meaning = model.classify(simplified)
    if meaning == "database":
        x = input("add/remove/view \n")
        if x.lower() == "add":
            add_data()
        elif x.lower() == "view":
            view()
        elif x.lower() == "remove":
            delete()
        continue
    else:
        if meaning == "login":
            flag = False
            engine.say("Enter the account you want to log in to")
            engine.runAndWait()
            r = sr.Recognizer()
            mic = sr.Microphone()
            with mic as source:
                r.adjust_for_ambient_noise(source,duration=2)
                print("Speak Now")
                audio = r.listen(source)
            print("Recognising Audio")
            account = r.recognize_google(audio)
            account = account.lower()
            doc = openpyxl.load_workbook("Database.xlsx")
            sheet = doc["database"]
            for i in range(3,sheet.max_row+1):
                if sheet["B"+str(i)].value == account:
                    flag = True
                    user_name = sheet["E"+str(i)].value
                    pwd = sheet["F"+str(i)].value
                    browser = webdriver.Chrome()
                    browser.get(sheet["G"+str(i)].value)
                    if sheet["H"+str(i)].value == "id":
                        usr = browser.find_element_by_id(user_name)
                        usr.send_keys(sheet["C"+str(i)].value)
                        if sheet["J"+str(i)].value == "y":
                            usr.send_keys(Keys.ENTER)
                            time.sleep(3)
                    elif sheet["H"+str(i)].value == "name":
                        usr = browser.find_element_by_name(user_name)
                        usr.send_keys(sheet["C"+str(i)].value)
                        if sheet["J"+str(i)].value == "y":
                            usr.send_keys(Keys.ENTER)
                            time.sleep(3)
                    if sheet["I"+str(i)].value == "name":
                        pwds = browser.find_element_by_name(pwd)
                        pwds.send_keys(sheet["D"+str(i)].value)
                        pwds.send_keys(Keys.ENTER)
                    elif sheet["I"+str(i)].value == "id":
                        pwds = browser.find_element_by_id(pwd)
                        pwds.send_keys(sheet["D"+str(i)].value)
                        pwds.send_keys(Keys.ENTER)
            if flag == False:
                engine.say("No such account in database")
                engine.runAndWait()
            continue
        elif meaning == "email":
            send_email()
        elif meaning == "search":
            search_net()
        elif meaning == "inbox":
            mail()
        elif meaning == "wikipedia":
            wiki()
        elif meaning == "quote":
            quote()
        elif meaning == "map":
            map_()
        elif meaning == "sad":
            sad()
        elif meaning == "happy":
            happy()
        else:
            pprint.pprint("Invalid Command")
     
