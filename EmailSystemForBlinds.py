#!/usr/bin/env python
# coding: utf-8

# In[23]:


import speech_recognition as sr
import smtplib
import pyaudio
import platform
import sys
from bs4 import BeautifulSoup
import email
import imaplib
from gtts import gTTS
import pyglet
import os, time
import playsound
import pyodbc 


# In[4]:


tts = gTTS(text="Project: Voice based Email for blind", lang='en')
ttsname=('name.mp3')
playsound.playsound("name.mp3")


# In[5]:


tts = gTTS(text="speak now", lang='en')
ttsname=("speak.mp3")


# In[6]:


tts = gTTS(text=" ok done", lang='en')
ttsname=("od.mp3")


# In[7]:



tts = gTTS(text="speak your email address ", lang='en')
ttsname=('email.mp3')

playsound.playsound("email.mp3")
playsound.playsound("speak.mp3")
r = sr.Recognizer()
with sr.Microphone() as source:
   audio=r.listen(source)
playsound.playsound("od.mp3")
email=r.recognize_google(audio)
email=email.replace(" ","")
print(email)
  


# In[8]:


tts = gTTS(text="the email address entered by you is ::"+email, lang='en')
ttsname=('myemail1.mp3')

playsound.playsound(ttsname)


# In[11]:



tts = gTTS(text="speak your password", lang='en')
ttsname=('password.mp3')

playsound.playsound("password.mp3")
playsound.playsound("speak.mp3")
r = sr.Recognizer()
with sr.Microphone() as source:
   audio=r.listen(source)
playsound.playsound("od.mp3")
password=r.recognize_google(audio)
password=password.replace(" ","")
password=password.lower();
print(password)
  


# In[12]:


login = os.getlogin
print ("You are logging trying from : "+email)


# In[14]:


tts = gTTS(text="option a. composed a mail.", lang='en')
ttsname=("aa.mp3")

playsound.playsound(ttsname)
print ("b. Check your inbox")
tts = gTTS(text="option b. Check your inbox", lang='en')
ttsname=("send.mp3")

playsound.playsound(ttsname)


# In[15]:


tts = gTTS(text="Your choice ", lang='en')
ttsname=("hello3.mp3")
playsound.playsound(ttsname)


# In[20]:


r = sr.Recognizer()
with sr.Microphone() as source:
    print ("Your choice:")
    audio=r.listen(source)
    print ("ok done!!")

try:
    text1=r.recognize_google(audio)
    print(text1)
    if text1=='a':
        tts = gTTS(text="you have choose option A", lang='en')
        ttsname=("qd.mp3")
    
        playsound.playsound("qd.mp3")
    elif text1=='b':
        tts = gTTS(text="you have choose option b ", lang='en')
        ttsname=("qb.mp3")
        
        playsound.playsound("qb.mp3")
    else:
        tts = gTTS(text="you have choose invalid option ", lang='en')
        ttsname=("qc.mp3")
      
        playsound.playsound("qc.mp3")
    
except sr.UnknownValueError:
    print("Google Speech Recognition could not understand audio.")
     
except sr.RequestError as e:
    print("Could not request results from Google Speech Recognition service; {0}".format(e)) 


# In[21]:


if text1 == 'a':
    tts = gTTS(text="speak receiver's email address ", lang='en')
    ttsname=('remail.mp3')
    playsound.playsound("remail.mp3")
    playsound.playsound("speak.mp3")
    r = sr.Recognizer()
    with sr.Microphone() as source:
       audio=r.listen(source)
    playsound.playsound("od.mp3")
    tomail=r.recognize_google(audio)
    tomail=tomail.replace(" ","")
    print(tomail)
    r = sr.Recognizer() #recognize
    with sr.Microphone() as source:
        tts = gTTS(text="enter your message body", lang='en')
        ttsname=("message.mp3")
      
        playsound.playsound("message.mp3")
        audio=r.listen(source)
        playsound.playsound("od.mp3")
    try:
        text2=r.recognize_google(audio)
        print ("You said : "+text2)
        msg = text2
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio.")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))    

    mail = smtplib.SMTP('smtp.gmail.com',587)    #host and port area
    mail.ehlo()  #Hostname to send for this command defaults to the FQDN of the local host.
    mail.starttls() #security connection
    mail.login('harshitrockslasec@gmail.com','harshit0209') #login part
    mail.sendmail('harshitrockslasec@gmail.com','samarth109.jyc@gmail.com',msg) #send part
    print ("Congrats! Your mail has send. ")
    tts = gTTS(text="Congrats! Your mail has send. ", lang='en')
    ttsname=("send12.mp3")
    playsound.playsound(ttsname)
    mail.close()   


# In[26]:


if text1 == 'b':
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-VTUCSFN\SQLEXPRESS01;'
                      'Database=mail;'
                      'Trusted_Connection=yes;')
    mail = imaplib.IMAP4_SSL('imap.gmail.com',993) #this is host and port area.... ssl security
    unm = ('harshitrockslasec@gmail.com')  #username
    psw = ('harshit0209')  #password
    mail.login(unm,psw)  #login
    stat, total = mail.select('Inbox')  #total number of mails in inbox
    print ("Number of mails in your inbox :"+str(total))
    tts = gTTS(text="Total mails are :"+str(total), lang='en') #voice out
    playsound.playsound("total.mp3")
    
#     music = pyglet.media.load(ttsname, streaming = False)
#     music.play()
#     time.sleep(music.duration)
#     os.remove(ttsname)
#     #unseen mails

#     music = pyglet.media.load(ttsname, streaming = False)
#     music.play()
#     time.sleep(music.duration)
#     os.remove(ttsname)
#     #search mails
    result, data = mail.uid('search',None, "ALL")
    inbox_item_list = data[0].split()
    new = inbox_item_list[-5]
    old = inbox_item_list[0]
    result2, email_data = mail.uid('fetch', new, '(RFC822)') #fetch
    raw_email = email_data[0][1].decode("utf-8") #decode
    email_message = email.message_from_string(raw_email)
    print ("From: "+email_message['From'])
    q=str(email_message['From'])
    w=str(email_message['Subject'])
    print ("Subject: "+str(email_message['Subject']))
    
    tts = gTTS(text="From: "+email_message['From']+" And Your subject: "+str(email_message['Subject']), lang='en')
    playsound.playsound("mail.mp3")
 
#     music = pyglet.media.load(ttsname, streaming = False)
#     music.play()
#     time.sleep(music.duration)
#     os.remove(ttsname)
#     #Body part of mails
    stat, total1 = mail.select('Inbox')
    stat, data1 = mail.fetch(total1[0], "(UID BODY[TEXT])")
    msg = data1[0][1]
    soup = BeautifulSoup(msg, "html.parser")
    txt = soup.get_text()
    print ("Body :"+txt)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO mail.dbo.datatable(sender_id,subject,body)VALUES('q','w','txt')")
    tts = gTTS(text="Body: "+txt, lang='en')
    #playsound.playsound("body.mp3")
    
#     music = pyglet.media.load(ttsname, streaming = False)
#     music.play()
#     time.sleep(music.duration)
#     os.remove(ttsname)
#     mail.close()
  




    cursor = conn.cursor()
    cursor.execute('SELECT * FROM mail.dbo.datatable')

    for row in cursor:
        print(row)
    mail.logout()


# In[ ]:





# In[ ]:




