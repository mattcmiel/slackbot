import xlrd
import xlsxwriter
import requests

#token for further faster
token = 'xoxp-642407649986-715987827012-1015550291508-c7630f088f8a84a357aa3d5f5b7d3e9a'


def post_message_to_slack(text, time, user, blocks = None):
    return requests.post('https://slack.com/api/reminders.add', {
        'token': token,
        'time':time,
        'text': text,
        'user':user,
        'blocks': json.dumps(blocks) if blocks else None
    }).json()

#First name
workbook = xlrd.open_workbook('masterlist.xlsx')
sheet = workbook.sheet_by_index(0)
myFirstName = []
for a in range (0,len(sheet.col_values(0))):
    value = sheet.cell_value(a,0)
    myFirstName.append(value)

#last name
workbook = xlrd.open_workbook('masterlist.xlsx')
sheet = workbook.sheet_by_index(0)
myLastName = []
for b in range (0,len(sheet.col_values(0))):
    value1 = sheet.cell_value(b,1)
    myLastName.append(value1)

#email
workbook = xlrd.open_workbook('masterlist.xlsx')
sheet = workbook.sheet_by_index(0)
myEmail = []
for c in range (0,len(sheet.col_values(0))):
    value2 = sheet.cell_value(c,2)
    myEmail.append(value2)

#company
workbook = xlrd.open_workbook('masterlist.xlsx')
sheet = workbook.sheet_by_index(0)
myCompany = []
for d in range (0,len(sheet.col_values(0))):
    value3 = sheet.cell_value(d,3)
    myCompany.append(value3)

#role
workbook = xlrd.open_workbook('masterlist.xlsx')
sheet = workbook.sheet_by_index(0)
myPosition = []
for e in range (0,len(sheet.col_values(0))):
    value4 = sheet.cell_value(e,4)
    myPosition.append(value4)


#times
workbook = xlrd.open_workbook('bottime.xlsx')
sheet = workbook.sheet_by_index(0)
times = []
for q in range (0,len(sheet.col_values(0))):
    value6 = sheet.cell_value(q,0)
    times.append(int(value6))



post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[0] + " " + myLastName[0] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/shiek-astin-bey-05190a16b/" + "\nEmail: " + myEmail[0] + "\nCompany: " + myCompany[0] + "\nRole: " + myPosition[0] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[0],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[1] + " " + myLastName[1] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/autumnschultz/" +"\nEmail: " + myEmail[1] + "\nCompany: " + myCompany[1] + "\nRole: " + myPosition[1] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[0],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[2] + " " + myLastName[2] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/rhonda-gilmore-a3163a11/" +"\nEmail: " + myEmail[2] + "\nCompany: " + myCompany[2] + "\nRole: " + myPosition[2] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[1],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[3] + " " + myLastName[3] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/grantkmartin/" +"\nEmail: " + myEmail[3] + "\nCompany: " + myCompany[3] + "\nRole: " + myPosition[3] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[1],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[4] + " " + myLastName[4] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/irina-litchfield-b50352131/" +"\nEmail: " + myEmail[4] + "\nCompany: " + myCompany[4] + "\nRole: " + myPosition[4] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[2],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[5] + " " + myLastName[5] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/jabalera/" +"\nEmail: " + myEmail[5] + "\nCompany: " + myCompany[5] + "\nRole: " + myPosition[5] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[2],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[6] + " " + myLastName[6] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/caleb-reynolds-36b86969/" +"\nEmail: " + myEmail[6] + "\nCompany: " + myCompany[6] + "\nRole: " + myPosition[6] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[3],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[7] + " " + myLastName[7] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/carlos-berrout-3925a8169/" +"\nEmail: " + myEmail[7] + "\nCompany: " + myCompany[7] + "\nRole: " + myPosition[7] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[3],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[8] + " " + myLastName[8] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/bklaydman/" +"\nEmail: " + myEmail[8] + "\nCompany: " + myCompany[8] + "\nRole: " + myPosition[8] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[4],'UKA0DAG6S')
post_message_to_slack("\nDaily Warm UPs Bot"+"\nName: " + myFirstName[9] + " " + myLastName[9] + "\nLinkedin Profile: " + "https://www.linkedin.com/in/sarah-eckstein-a4a402a/" +"\nEmail: " + myEmail[9] + "\nCompany: " + myCompany[9] + "\nRole: " + myPosition[9] +"\nhttps://www.linkedin.com/messaging/thread/new/", times[4],'UKA0DAG6S')


print("done")
