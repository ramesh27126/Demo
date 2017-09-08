from xlrd import open_workbook
import smtplib
from email.mime.multipart import MIMEMultipart


def email(filepath):
    wb=open_workbook(filepath)
    sent_mail = []
    mylist={}

    for sheet in wb.sheets():
        for row in range(1,sheet.nrows):
            userid=sheet.cell_value(row,3)
            mylist[userid]=[]
            mylist[userid].append('{}'.format(sheet.cell_value(row,4)))
            mylist[userid].append('{}'.format(sheet.cell_value(row,5)))


        for value in mylist:
            symbol = 1
            for row in range(0, sheet.nrows):
                if value == sheet.cell_value(row,3):
                    question=sheet.cell_value(row,0)
                    mylist[value].append('{}.{}'.format(symbol, question))
                    symbol+=1

    for id in mylist:
        user_id = id

        user_name = mylist[id][1]
        fromaddr = 'casper@techmahindra.com'
        toaddrs = mylist[id][0]
        message = MIMEMultipart()
        message['subject']='casperhelpdesk'
        message['From']=fromaddr
        message['To']=toaddrs
        message.preamble='casperhelpdesk'
        msg= (
        "From:{}\n"
        "To:{}\n\n"
        "Hi {},\n\nHere are the questions that we have updated in CASPER based on your request\n ".format(fromaddr, toaddrs, user_name))
        for i in range(2, len(mylist[id])):
            msg = ("{}\n\t{}".format(msg, mylist[id][i]))
        signature = ("WITH REGARDS\nCASPER HELPDESK")
        message['boby'] = (msg + "\n\n{}".format(signature))
        sent_mail.append("{} : Email Sent Successfully".format(user_id))
        # server = smtplib.SMTP('PUNEXCHMBX001.Techmahindra.com')
        # server.set_debuglevel(1)
        # server.sendmail(fromaddr, toaddrs, message.as_string())
        # server.quit()

        print(user_id, user_name,fromaddr, toaddrs)
        print(msg)



email(filepath='C:\\Users\\RR00484222\\PycharmProjects\\RAMESH\\excel_package\\TLS_Digital.xlsx')



