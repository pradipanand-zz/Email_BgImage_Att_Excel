# Email_BgImage_Att_Excel
'''
Hardly you will be able to find about how to have background image in outlook meail and also excel sheet as attachment, 
bottom line HTML + excelsheet, this code will work in both case

server - Login to outlook and send details to this module
to - send to persone email id
from_email - from which email you want to send
name - intended person name
Employee_Number - If need
random_group_team- excel (xlsx) file path (attachement)
random_group_image_path - background path
team_name - Team Name 

***** I have created this small def to send email to all employee with the background image and team name, 
if you want to complete script, please mail pradipanand@gmail.com
'''
def Email_Send(server, to, from_email, name, Employee_Number, random_group_team, random_group_image_path, team_name):

    # cc = "anotherperson@gmail.com,someone@yahoo.com"
    # bcc = "pradipanand@gmail.com"

    msg             = MIMEMultipart('related')
    # msg             = MIMEMultipart('alternative')
    msg['Subject']  = "%s" % ('"Funtastic Five"')
    msg['From']     = from_email
    msg['To']       = to
    # msg['bcc']      = 'pradipanand@gmail.com'

    html = """\
    <html>
      <head></head>
        <body>
            <p style="font-size:15px;">Congratulations <b> """+ name +""" </b>& welcome to <b>"""+ team_name +"""" !</b> Also, please find attached the list of your fellow team members.</p>
            <br>
            <br>
            <table align="center" style="margin: 10px auto;">
             <tr>
              <td>
                <img src="cid:image2" width="1000" height="600" alt="Team Name">
              </td>
            </tr>
            </table>
        </body>
    </html>
    """
    part2 = MIMEText(html, 'html', 'utf-8')
    msg.attach(part2)

    fp = open(random_group_image_path, 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', '<image2>')
    msg.attach(msgImage)

    attachment = open(random_group_team, 'rb')
    xlsx = MIMEBase('application','vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    xlsx.set_payload(attachment.read())
    encoders.encode_base64(xlsx)
    xlsx.add_header('Content-Disposition', 'attachment', filename="Meredith India Teams.xlsx")
    msg.attach(xlsx)
