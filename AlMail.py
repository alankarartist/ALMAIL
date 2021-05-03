import smtplib
from smtplib import SMTPException
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path
import os
from tkinter import*
from tkinter import font
import pyttsx3
import sys

cwd = os.path.dirname(os.path.realpath(__file__))

class AlMail:
    def __init__(self,mailType):
        if mailType.lower()=='gmail':
            root= Tk(className=" ALGMAIL " )
            root.geometry("500x700+1410+315")
            root.config(bg="#e22b2d")
            color='#e22b2d' 
        elif mailType.lower()=='outlook' or mailType.lower()=='hotmail' or mailType.lower()=='live':
            root= Tk(className=" ALMICROSOFT " )
            root.geometry("500x700+1410+315")
            root.config(bg="#035aaa")
            color='#035aaa'      
        root.resizable(0,0)
        root.iconbitmap(os.path.join(cwd+'\\UI\\icons', 'almail.ico'))
        
        def speak(audio):
            engine = pyttsx3.init('sapi5')
            voices = engine.getProperty('voices')
            engine.setProperty('voice', voices[0].id)
            engine.say(audio)
            engine.runAndWait()

        def sendemail():         
            usermail = userEmail.get()
            tomail = toEmail.get()  
            ccmail = ccEmail.get()
            bccmail = bccEmail.get()
            password = passWord.get()
            subject = subj.get()
            mainMessage = body.get('1.0', 'end-1c')
            filename = fileName.get()
            msg = MIMEMultipart()
            msg['From'] = usermail
            to=[]
            if tomail:
                if ',' not in tomail:
                    to.append(tomail)
                    msg['To'] = tomail
                if ',' in tomail:
                    to = tomail.split(',')
                    msg['To'] = ', '.join(to)
            cc=[]
            if ccmail:
                if ',' not in ccmail:
                    cc.append(ccmail)
                    msg['Cc'] = ccmail
                if ',' in ccmail:
                    cc = ccmail.split(',')
                    msg['Cc'] = ', '.join(cc)
            bcc=[]
            if bccmail:
                if ',' not in bccmail:
                    bcc.append(bccmail)
                    msg['Bcc'] = bccmail
                if ',' in bccmail:
                    bcc = bccmail.split(',')
                    msg['Bcc'] = ', '.join(bcc)
            msg['Subject'] = subject
            msg.attach(MIMEText(mainMessage, 'plain'))
            if filename:
                if ',' not in filename:
                    filepath = os.path.join(os.path.abspath(cwd)+'\AlMail\Attachments',filename)
                    attachname = os.path.basename(filepath)
                    attachment = open(filepath, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', "attachment; filename= %s" % attachname)
                    msg.attach(part)
                if ',' in filename:
                    filename = filename.split(',')
                    for f in filename:
                        filepath = os.path.join(os.path.abspath(cwd)+'\AlMail\Attachments',f)
                        attachname = os.path.basename(filepath)
                        attachment = open(filepath, "rb")
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', "attachment; filename= %s" % attachname)
                        msg.attach(part)
            try:
                text.delete(1.0, END)
                if mailType.lower()=='gmail':
                    server=smtplib.SMTP('smtp.gmail.com','587')
                elif mailType.lower()=='outlook' or mailType.lower()=='hotmail' or mailType.lower()=='live':
                    debuglevel = True
                    server=smtplib.SMTP('smtp.office365.com','587')
                    server.set_debuglevel(debuglevel)
                server.ehlo()
                server.starttls()
                server.login(usermail, password)
                textmsg = msg.as_string()
                server.sendmail(usermail, (to+cc+bcc), textmsg)
                text.insert(1.0, 'Mail sent. ')
                speak('Mail sent. ')
            except(smtplib.SMTPException,ConnectionRefusedError,OSError):
                text.insert(1.0, 'Mail not sent. ')
                speak('Mail not sent. ')
            finally:
                server.quit()
        appHighlightFont = font.Font(family='sans-serif', size=12, weight='bold')
        textHighlightFont = font.Font(family='Segoe UI', size=12)

        #user mail
        userEmail = Label(root, text="USERNAME")
        userEmail.pack()
        userEmail.config(bg=color,fg="white",font=textHighlightFont)
        userEmail = Entry(root, highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        userEmail.pack(fill=X)

        #to email
        toEmail = Label(root, text="TO")
        toEmail.pack( )
        toEmail.config(bg=color,fg="white",font=textHighlightFont)
        toEmail = Entry(root, highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        toEmail.pack(fill=X)

        #cc email
        ccEmail = Label(root, text="CC")
        ccEmail.pack( )
        ccEmail.config(bg=color,fg="white",font=textHighlightFont)
        ccEmail = Entry(root, highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        ccEmail.pack(fill=X)

        #bcc email
        bccEmail = Label(root, text="BCC")
        bccEmail.pack( )
        bccEmail.config(bg=color,fg="white",font=textHighlightFont)
        bccEmail = Entry(root, highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        bccEmail.pack(fill=X)

        #subject line
        subj= Label(root, text="SUBJECT")
        subj.pack( )
        subj.config(bg=color,fg="white",font=textHighlightFont)
        subj = Entry(root, highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        subj.pack(fill=X)

        #Body of the message
        body = Text(root, font="sans-serif",  relief=SUNKEN , highlightbackground=color, highlightcolor=color, highlightthickness=5, bd=0)
        body.config(bg="black", fg='white', height=10, font=appHighlightFont)
        body.pack(fill=BOTH, expand=True)

        #Attachments to send
        fileName = Label(root, text="ATTACHMENT")
        fileName.pack( )
        fileName.config(bg=color,fg="white",font=textHighlightFont)
        fileName = Entry(root, highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        fileName.pack(fill=X)

        #passWord widget
        passWord = Label(root, text="PASSWORD")
        passWord.pack()
        passWord.config(bg=color,fg="white",font=textHighlightFont)
        passWord= Entry(root, show='*', highlightbackground=color, highlightcolor=color, highlightthickness=3, bd=0,font=appHighlightFont)
        passWord.pack(fill=X)

        #submit button
        submitMail = Button(root, borderwidth=0, highlightthickness=5, text="SEND MAIL", command=sendemail)
        submitMail.config(bg=color,fg="white",font=textHighlightFont)
        submitMail.pack(fill=X)

        #feed back
        text = Text(root, font="sans-serif",  relief=SUNKEN , highlightbackground=color, highlightcolor=color, highlightthickness=5, bd=0)
        text.config(bg="black", fg='white', height=2, font=appHighlightFont)
        text.pack(fill=BOTH, expand=True)

        root.mainloop()

if __name__=='__main__':
    AlMail(sys.argv[1]) 