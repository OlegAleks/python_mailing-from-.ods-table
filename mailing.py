import datetime
from datetime import date
from pyexcel_ods import get_data
from pyexcel_ods import save_data
from collections import OrderedDict
import json
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
def send_mail(fromaddr, toaddr, msg):
	server.sendmail("olegusfox@gmail.com", toaddr, msg)
def deunicodify(stroke):
	if isinstance(stroke, list):
		new_list = []
		for i in stroke:
			i = deunicodify(i)
			new_list.append(i)
		stroke = new_list
	if isinstance(stroke, unicode):
		stroke = stroke.encode('utf-8')
	return stroke
def deunicodify_hook(pairs):
	new_pairs = []
	for key, value in pairs:
		key = deunicodify(key)
		value = deunicodify(value)
		new_pairs.append((key, value))
	return dict(new_pairs)
def check_date(row):
	today = date.today()
	last_email = row[4]
	if last_email != "":
		last_email = datetime.datetime.strptime(last_email.replace("-"," "), "%Y %m %d")
		last_email = last_email.date()
	else:
		last_email = today-datetime.timedelta(days=2)
# delta - is how many days must spent to be able to send new mail to this adress
	delta = datetime.timedelta(days=1)
	check = 0
	new_row = row
	if today>last_email+delta:
		new_row[4] = today.strftime("%Y-%m-%d")
		check = 1
	return (new_row, check)
#path to .ods file with data base of your contacts
file = "D:\table.ods"
data = get_data(file)
j_str = json.dumps(data)
dic = json.loads(j_str, object_pairs_hook=deunicodify_hook)
info = dic['Sheet1']
new_info = []
for row in info:
	L = len(row)
	new_row = []
	if L>0 and L<5:
		for i in range(0, 5):
			try:
				cell = row[i]
				new_row.append(cell)
			except:
				new_row.append("")
	elif L==5: new_row = row
	if L!=0: new_info.append(new_row)
info = new_info
# connect to gmail SMTP and login
server = smtplib.SMTP('smtp.gmail.com', 587)
server.ehlo()
server.starttls()
#your App-password on google
app_pass = 'xxxxxxxxxxxxxxxxxx' 
server.login("olegusfox@gmail.com", app_pass)
fromaddr = "olegusfox@gmail.com"
i = 0
new_info = []
for row in info:
	if i ==0:
		i+=1
		new_info.append(row)
		continue
	check = check_date(row)
	if check[1]:
		company = row[0]
		person = row[1]
		toaddr = row[2]
		artist = "Oleg Alekseev"
		gender = row[3]
		if gender != "":
			solutation = "Dear " + gender + " %s\n\n"%(person)
		else:
			solutation = "Dear %s\n\n"%(person)
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Subject'] = 'FX TD Application from %s'%(artist)
		contact_info = "%s\nSankt-Peterburg, Russian Federation\nolegusfox@gmail.com\n+7 904 639 17 32\n\n"%(artist)
		part1 = "My name is %s.\nI would like to express my interset in your position for a FX TD at %s.\n\n"%(artist,company)
		part2 = "I have extensive knowledge of the technology used in creating visual effects and I would appreciate an opportunity to put my skills and training to work for %s.\n"%(company)
		part3 = "My work experiance in CG starts in 2011 year at post-production studio named 'Algous'\nNow I'm working for 'Melnitsa animation studio' as Houdini FX TD.\n\n"
		SH = "Please find my showreel at the following link: https://www.youtube.com/watch?v=K4xtXj6CRbY\n"
		LN = "LinkedIn: https://www.linkedin.com/in/oleg-alekseev-53a2a783/\n"
		CV = "CV: http://olegusfox.wixsite.com/alexeev-oleg/cv\n\n"
		END = "Respectfully,\n%s\nFX TD"%(artist)
		body = contact_info+solutation+part1+part2+part3+SH+LN+CV+END
		msg.attach(MIMEText(body, 'plain'))
		txt = msg.as_string()
		send_mail(fromaddr, toaddr, txt)
	new_info.append(check[0])
server.quit()
dic['Sheet1'] = new_info
data_to_save = OrderedDict()
data_to_save.update(dic)
new_file = file
save_data(new_file, dic)
