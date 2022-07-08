import logging
import json
import pymysql
import datetime
import smtplib
import xlsxwriter
#################
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
#################
from rich.console import Console
from rich.table import Table

#open json config file
try:
    with open('config.json', 'r') as cnf:
        config_json = cnf.read()
    config = json.loads(config_json)
except:
    raise Exception('error config')

#main init
test_mode = config['main']['test_mode']

if test_mode == True:
    scan_all = config['mysql']['table_scan_all_for_test_mode']
    workplace_data = config['mysql']['table_workplace_data_for_test_mode']
else:
    scan_all = config['mysql']['table_scan_all']
    workplace_data = config['mysql']['table_workplace_data']


#date_now = datetime.date.today().strftime('%Y-%m-%d')
date_now = '2022-07-06'
time_now = datetime.datetime.now().time().strftime('%H:%M:%S')


#mysql connect
try:
    db_host = config['mysql']['host_to_mysql']
    db_user = config['mysql']['user_to_mysql']
    db_password = config['mysql']['password_to_mysql']
    db_database = config['mysql']['database_name_to_mysql']
    db_connection = pymysql.connect(
        host=db_host,
        user=db_user,
        password=db_password,
        database=db_database)
    db_cursor = db_connection.cursor()
except:
    print('dberror')
    #logging.error('error connect to mysql')


##########################
console = Console()
file_to_output = open('report.txt', 'w', encoding='ascii')
file_console = Console(file=file_to_output)

db_cursor.execute("SELECT COUNT(numer) FROM "+scan_all+" WHERE time<='22:10:00' and date='"+date_now+"' and open_close=0")
rows = db_cursor.fetchall()
file_console.print('count scan all = '+str(rows[0][0]))

db_cursor.execute("SELECT COUNT(numer) FROM "+workplace_data+" WHERE time<='22:10:00' and date='"+date_now+"' and open_close=0")
rows = db_cursor.fetchall()
file_console.print('count open in workplace = '+str(rows[0][0]))

db_cursor.execute("SELECT COUNT(numer) FROM "+workplace_data+" WHERE time<='22:10:00' and date='"+date_now+"' and open_close=1")
rows = db_cursor.fetchall()
file_console.print('count close in workplace = '+str(rows[0][0]))

db_cursor.execute("SELECT COUNT(numer) FROM "+scan_all+" WHERE time<='22:10:00' and date='"+date_now+"' and open_close=1")
rows = db_cursor.fetchall()
file_console.print('count close in scan all = '+str(rows[0][0]))

db_cursor.execute("SELECT * FROM "+workplace_data+" WHERE time<='22:10:00' and date='"+date_now+"' and open_close=1")
rows = db_cursor.fetchall()

#xlsx init
workbook = xlsxwriter.Workbook('report.xlsx')
##
bold = workbook.add_format({'bold': True})
bold.set_bg_color('green')
#close_in_workplace
worksheet_close_in_workplace = workbook.add_worksheet('close in workplace')
worksheet_close_in_workplace.set_column(0, 0, 15)
worksheet_close_in_workplace.set_column(6, 6, 13)

worksheet_close_in_workplace.write(0, 0, 'numer', bold)
worksheet_close_in_workplace.write(0, 1, 'time', bold)
worksheet_close_in_workplace.write(0, 2, 'position', bold)
worksheet_close_in_workplace.write(0, 3, 'count', bold)
worksheet_close_in_workplace.write(0, 4, 'type', bold)
worksheet_close_in_workplace.write(0, 5, 'worker', bold)
worksheet_close_in_workplace.write(0, 6, 'time summary', bold)

i_in_while = 0
while i_in_while < len(rows):
    worksheet_close_in_workplace.write(i_in_while+1, 0, rows[i_in_while][0])
    worksheet_close_in_workplace.write(i_in_while+1, 1, str(rows[i_in_while][3]))
    worksheet_close_in_workplace.write(i_in_while+1, 2, rows[i_in_while][4])
    worksheet_close_in_workplace.write(i_in_while+1, 3, rows[i_in_while][5])
    worksheet_close_in_workplace.write(i_in_while+1, 4, str(rows[i_in_while][6]))
    worksheet_close_in_workplace.write(i_in_while+1, 5, rows[i_in_while][7])
    worksheet_close_in_workplace.write(i_in_while+1, 6, str(rows[i_in_while][8]))
    i_in_while += 1

#close_in_scan_all
worksheet_close_in_scan_all = workbook.add_worksheet('close in scan all')
worksheet_close_in_scan_all.set_column(0, 0, 20)
worksheet_close_in_scan_all.write(0, 0, 'numer', bold)
worksheet_close_in_scan_all.write(0, 1, 'time', bold)

##
db_cursor.execute("SELECT * FROM "+scan_all+" WHERE time<='22:10:00' and date='"+date_now+"' and open_close=1")
rows = db_cursor.fetchall()

i_in_while = 0
while i_in_while < len(rows):
    worksheet_close_in_scan_all.write(i_in_while+1, 0, rows[i_in_while][0])
    worksheet_close_in_scan_all.write(i_in_while+1, 1, str(rows[i_in_while][2]))
    i_in_while += 1

workbook.close()
##
file_to_output.close()

#mail
file_to_output = open('report.txt', 'r')
str_to_mail = file_to_output.read()
file_to_output.close()
##
msg = MIMEMultipart()
msg['From'] = config['email']['addr_from']
msg['To'] = config['email']['addr_to']
##
email = smtplib.SMTP('wn30.webd.pl', 587)
email.starttls()
email.login(config['email']['addr_from'], config['email']['password'])
##
msg.attach(MIMEText(str_to_mail, 'plain'))#<--
##
part = MIMEBase('application', "octet-stream")
part.set_payload(open("report.xlsx", "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="repot.xlsx"')
msg.attach(part)
##
msg['Subject'] = 'subject'
email.send_message(msg)
##
email.quit()

db_connection.close()






