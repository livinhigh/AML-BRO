# https://myaccount.google.com/u/1/lesssecureapps?pli=1&pageId=none
#and allow imap access on gmail settings
#user has to go first to the link they are logging in with and allow less secure apps
import datetime
import smtplib
import time
import base64
import imaplib
import webbrowser
import os.path
from os import path
import email
import pickle
from win10toast import ToastNotifier
todtom=raw_input("TODAY OR TOMORROQ? (to/tm):")

today = datetime.date.today()
tommorow = today + datetime.timedelta(days = 1)
if todtom.lower()=="tm":
   strtomorrow= tommorow.strftime("%A")+", "+tommorow.strftime("%B")+" "+tommorow.strftime("%d")+", 2020"
else:
   tommorow=today
   strtomorrow= tommorow.strftime("%A")+", "+tommorow.strftime("%B")+" "+tommorow.strftime("%d")+", 2020"
#1-user id and 2 - password
ORG_EMAIL   = "@btech.christuniversity.in"
if os.path.exists("userlogin.pickle") :
   pickle_in=open("userlogin.pickle","rb")
   user_dict=pickle.load(pickle_in)
   FROM_EMAIL=user_dict[1]
   FROM_PWD=user_dict[2]
   pickle_in.close()
   print "Logging in as "+FROM_EMAIL
   
else:
   FROM_EMAIL  = raw_input("Enter the user id without @btech.christuniversity.in :")
   FROM_EMAIL = FROM_EMAIL + ORG_EMAIL
   FROM_PWD    =raw_input("Enter password of " + FROM_EMAIL + " : ")
   pickle_out=open("userlogin.pickle","wb")
   user_dict={1:FROM_EMAIL,2:FROM_PWD}
   pickle.dump(user_dict,pickle_out)
   pickle_out.close()

SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993
"""
f1=open("meeting.pickle","wb")
meeting_details={0:"not updated",1:strtomorrow,2:"DAA",3:"NO EMAIL",4:""}
pickle.dump(meeting_details,f1)
meeting_details={0:"not updated",1:strtomorrow,2:"Internet and Web Programming",3:"NO EMAIL",4:""}
pickle.dump(meeting_details,f1)
meeting_details={0:"not updated",1:strtomorrow,2:"Software Engineering",3:"NO EMAIL",4:""}
pickle.dump(meeting_details,f1)
meeting_details={0:"not updated",1:strtomorrow,2:"Internet of Things",3:"NO EMAIL",4:""}
pickle.dump(meeting_details,f1)
meeting_details={0:"not updated",1:strtomorrow,2:"CONA",3:"NO EMAIL",4:""}
pickle.dump(meeting_details,f1)
meeting_details={0:"not updated",1:strtomorrow,2:"FLAT",3:"NO EMAIL",4:""}
pickle.dump(meeting_details,f1)
f1.close()
"""
# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
def read_email_from_gmail():
   #print "hello"
   try:
      mail = imaplib.IMAP4_SSL(SMTP_SERVER)
      mail.login(FROM_EMAIL,FROM_PWD)
      mail.select('inbox')

      type, data = mail.search(None, '(FROM "messenger@webex.com")')
      mail_ids = data[0]

      id_list = mail_ids.split()   
      first_email_id = int(id_list[0])
      latest_email_id = int(id_list[-1])
      print "fetched all mails"
      for i in range(latest_email_id,int(id_list[-20]), -1):
         print "\n-------------------\n"
         print "fetched " + str(i) + " mail id"
         typ, data = mail.fetch(i, '(RFC822)' )
         for response_part in data:
            print "check if conditions match"
            if isinstance(response_part, tuple):
                 msg = email.message_from_string(response_part[1])
                 email_subject = msg['subject']
                 email_from = msg['from']
                 #print('From : ' + email_from + '\n')
                 #print('Subject : ' + email_subject + '\n')
                 
                 body=get_body(msg)
                 #print strtomorrow
                 print "extracted message body"
                 if body.find(strtomorrow)!=-1:
                    print "matched conditions !"
                    #print body
                    if body.find(" am ")!=-1:
                       c_time=body[body.find(" am ")-5:body.find(" am ")]+" am"
                    else:
                       c_time=body[body.find(" pm ")-5:body.find(" pm ")]+" pm"
                    #print c_time

                    if body.find("REGISTER")!=-1:
                       status="NOT REGISTERED"
                       link=body[body.find("REGISTER")+9:(body.find("REGISTER")+100)]
                       #print link
                    elif body.find("When it's time, join the meeting.")!=-1:
                       status="REGISTERED. READY TO JOIN"
                       link=body[(body.find("When it's time, join the meeting.")+33):(body.find("When it's time, join the meeting.")+126)]
                    #print link
                    lowerbody=body.lower()
                    if lowerbody.find("software")!=-1:
                       subject = "Software Engineering"
                    elif lowerbody.find("internet and")!=1:
                       subject= "Internet and Web Programming"
                    elif lowerbody.find("daa")!=-1:
                       subject="DAA"
                    elif lowerbody.find("formal")!=1:
                       subject= "FLAT"
                    elif lowerbody.find("numerical ")!=-1:
                       subject="CONA"
                    elif lowerbody.find("iot")!=-1 or lowerbody.find("internet of things")!=-1:
                       subject="Internet of Things"
                    print subject
                    print "determined subject"
                    #print "check if file exists and update it"
                  
##      1.Design and Analysis of Algorithm
##         2.Internet Web Programming
##         3.Software Engineering
##         4.Internet of Things
##         5.CONA
##         6.Formal Language and Automata Theory- Elective

     
                    print "check if file exists and update it"
                    if os.path.exists("meeting.pickle") :
                       
                       print "found meeting pickle"
                       #changeflag=0
                       f1=open("meeting.pickle",'rb')
                       print "opened meeting.pickle"
                       while True:
                          try:
                             meeting_details=pickle.load(f1)
                             #print meeting_details 
                             if meeting_details[1]!=strtomorrow:
                                meeting_details[1]=strtomorrow
                                meeting_details[0]="not updated"
                                meeting_details[3]="NO EMAIL"
                                meeting_details[4]=""
                                print " inserted default values"
                             if meeting_details[2]==subject:
                                if meeting_details[3]!="REGISTERED. READY TO JOIN":
                                    if status=="REGISTERED. READY TO JOIN":
                                       meeting_details[3]=status
                                       meeting_details[4]=link
                                    else :
                                       meeting_details[3]=status
                                       meeting_details[4]=link
                                    meeting_details[0]=c_time   
                             f2=open("temp.pickle","ab")
                             print "opened temp file"
                             pickle.dump(meeting_details,f2)
                             print "info dumped into temp"
                             f2.close()
                          except:
                             f1.close()             
                             os.remove("meeting.pickle")
                             print "successfuly deleted"
                             os.rename('temp.pickle','meeting.pickle')
                             print "successfully renamed"
                             break
                    else:
                       print "meeting pickle missing"
                       f1=open("meeting.pickle","wb")
                       print "creating new meeting"
                       meeting_details={0:"not updated",1:strtomorrow,2:"DAA",3:"NO EMAIL",4:""}
                       pickle.dump(meeting_details,f1)
                       meeting_details={0:"not updated",1:strtomorrow,2:"Internet and Web Programming",3:"NO EMAIL",4:""}
                       pickle.dump(meeting_details,f1)
                       meeting_details={0:"not updated",1:strtomorrow,2:"Software Engineering",3:"NO EMAIL",4:""}
                       pickle.dump(meeting_details,f1)
                       meeting_details={0:"not updated",1:strtomorrow,2:"Internet of Things",3:"NO EMAIL",4:""}
                       pickle.dump(meeting_details,f1)
                       meeting_details={0:"not updated",1:strtomorrow,2:"CONA",3:"NO EMAIL",4:""}
                       pickle.dump(meeting_details,f1)
                       meeting_details={0:"not updated",1:strtomorrow,2:"FLAT",3:"NO EMAIL",4:""}
                       pickle.dump(meeting_details,f1)
                       f1.close()
                       print " inserted default values"
                       #changeflag=0
                       f1=open("meeting.pickle",'rb')
                       print "opened meeting.pickle"
                       while True:
                          try:
                             meeting_details=pickle.load(f1)
                             if meeting_details[1]!=strtomorrow:
                                meeting_details[1]=strtomorrow
                                meeting_details[0]="not updated"
                                meeting_details[3]="NO EMAIL"
                                meeting_details[4]=""
                                print " inserted default values"
                             if meeting_details[2]==subject:
                                if meeting_details[3]!="REGISTERED. READY TO JOIN":
                                    if status=="REGISTERED. READY TO JOIN":
                                       meeting_details[3]=status
                                       meeting_details[4]=link
                                    else :
                                       meeting_details[3]=status
                                       meeting_details[4]=link
                                    meeting_details[0]=c_time   
                             f2=open("temp.pickle","ab")
                             print "opened temp file"
                             pickle.dump(meeting_details,f2)
                             print "info dumped into temp"
                             f2.close()
                          except:
                             f1.close()             
                             os.remove('meeting.pickle')
                             print "original file removed"
                             os.rename('temp.pickle','meeting.pickle')
                             print "temp file renamed successfully"
                             break
                 
                    #meeting_details={0:c_time.strip("\n"),1:strtomorrow,2:subject,3:status,4:link.strip("\n")}
                    #print meeting_details
                 else:
                    print "not matched condition"
                    
                    
   except Exception, e:
      print str(e)
   updatehtml()   

def showtoast(message):
    toaster = ToastNotifier()
    toaster.show_toast(message)
    toaster.show_toast("Example two",
                   "This notification is in it's own thread!",
                   icon_path=None,
                   duration=5,
                   threaded=True)
    # Wait for threaded notification to finish
    while toaster.notification_active(): time.sleep(0.1)
def openbrowser(link):
   webbrowser.open(link,new=0,autoraise=True)
def updatehtml():
   f1=open("class list.html","w")
   if os.path.exists("meeting.pickle"):
      f2=open("meeting.pickle","rb")
      print "file found successfully"
      
   else:
      print "meeting pickle missing"
      f3=open("meeting.pickle","wb")
      print "creating new meeting pickle"
      meeting_details={0:"not updated",1:strtomorrow,2:"DAA",3:"NO EMAIL",4:""}
      pickle.dump(meeting_details,f3)
      meeting_details={0:"not updated",1:strtomorrow,2:"Internet and Web Programming",3:"NO EMAIL",4:""}
      pickle.dump(meeting_details,f3)
      meeting_details={0:"not updated",1:strtomorrow,2:"Software Engineering",3:"NO EMAIL",4:""}
      pickle.dump(meeting_details,f3)
      meeting_details={0:"not updated",1:strtomorrow,2:"Internet of Things",3:"NO EMAIL",4:""}
      pickle.dump(meeting_details,f3)
      meeting_details={0:"not updated",1:strtomorrow,2:"CONA",3:"NO EMAIL",4:""}
      pickle.dump(meeting_details,f3)
      meeting_details={0:"not updated",1:strtomorrow,2:"FLAT",3:"NO EMAIL",4:""}
      pickle.dump(meeting_details,f3)
      f3.close()
      print "default values info dumped"
   f2=open("meeting.pickle","rb")
   md1={}
   #print md1
   i=-1
   while True:
      try:
         i+=1
         update_md=pickle.load(f2)
         #print update_md
         md1[i]=update_md
      except:
         f2.close()
         f1.write('''<!DOCTYPE html><html><head><title>HTML Document</title><style>
         .button {
           border: none;
           color: white;
           padding: 16px 32px;
           text-align: center;
           text-decoration: none;
           display: inline-block;
           font-size: 16px;
           margin: 4px 2px;
           transition-duration: 0.4s;
           cursor: pointer;
         }

         .button1 {
           background-color: white;
           color: black;
           border: 2px solid #4CAF50;
         }

         .button1:hover {
           background-color: #4CAF50;
           color: white;
         }

         .button2 {
           background-color: white;
           color: black;
           border: 2px solid #008CBA;
         }

         .button2:hover {
           background-color: #008CBA;
           color: white;
         }

         </style>
         </head>
         <body style="text-align:center;">
         <h1 style="text-align:center;">Online Class Managemet<br></h1><h3 style="text-align:center;">Date : ''')
         f1.write(md1[0][1] +''' </h2><h2 style="text-align:center;">''')
         f1.write(md1[0][2] +'''<br></h1><p style="text-align:center;">at ''')
         f1.write(md1[0][0] +''' Status : '''+md1[0][3] )
         f1.write(''' </p><button class="button button1" onclick="window.location.href=\'''')
         f1.write(md1[0][4] +'''';">Join/Register</button><h2 style="text-align:center;">''')
         f1.write(md1[1][2] +''' <br></h1> <p style="text-align:center;">at '''+md1[1][0] +''' Status : '''+md1[1][3] )
         f1.write(''' </p><button class="button button1" onclick="window.location.href=\''''+md1[1][4] +''''';">Join/Register</button><h2 style="text-align:center;">'''+md1[2][2])
         f1.write(''' <br></h1><p style="text-align:center;">at '''+md1[2][0] +''' Status : '''+md1[2][3] +''' </p><button class="button button1" onclick="window.location.href=\''''+md1[2][4] )
         f1.write('''';">Join/Register</button><h2 style="text-align:center;">'''+md1[3][2] +'''<br></h1> <p style="text-align:center;">at '''+md1[3][0] +''' Status : '''+md1[3][3] )
         f1.write(''' </p><button class="button button1" onclick="window.location.href=\''''+md1[3][4] +'''';">Join/Register</button><h2 style="text-align:center;">'''+md1[4][2] )
         f1.write(''' <br></h1> <p style="text-align:center;">at '''+md1[4][0] +''' Status : '''+md1[4][3] +''' </p><button class="button button1" onclick="window.location.href=\''''+md1[4][4] )
         f1.write('''';">Join/Register</button> <h2 style="text-align:center;">'''+md1[5][2] +''' <br></h1> <p style="text-align:center;">at '''+md1[5][0] +''' Status : '''+md1[5][3] )
         f1.write(''' </p><button class="button button1" onclick="window.location.href=\''''+md1[5][4] +'''';">Join/Register</button>      </body>     </html>   ''')
         
         f1.close()
         break
   
read_email_from_gmail()
#updatehtml()
#openbrowser('https://myaccount.google.com/u/1/lesssecureapps?pli=1&pageId=none')
#showtoast("hellooo")
