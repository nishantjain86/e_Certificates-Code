from google.colab import auth
auth.authenticate_user()
import gspread
import pandas as pd
from oauth2client.client import GoogleCredentials
gc = gspread.authorize(GoogleCredentials.get_application_default())

import cv2
import matplotlib.pyplot as plt
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import matplotlib.pyplot as plt

# Put Email Subject
subject = "Certificate for  ISPCC 2021" 

# Put Email Body
body = "Dear Participant,\n        Thank You for presenting a paper in ISPCC-2021 organized by Department of Electronics and Communication Engineering, Jaypee University of Information Technology (JUIT), Waknaghat, Solan, H.P., INDIA."
body=body +"\n\nRegards \nOrganizing Committee \nISPCC 2021 \nDepartment of ECE\nJUIT, Waknaghat, Solan, HP, India"

# Put Sender Email ID and Password
sender_email = "nishantjain86@gmail.com"
password = "absdfghjkkklcvbn" 

# Put the text with which you want to send the file/Certificate.
filename1="ISPCC-2021,JUIT.jpg"

# Put the URL of the spread sheet 
wb = gc.open_by_url('https://docs.google.com/spreadsheets/d/1mRhotki-RiN_jQvD71NBgisvj-NkeD_u33xE-QIMaaU/edit#gid=0') 

# Put the name of the sheet that contains the List of certificates 
sheet = wb.worksheet('a')

data = sheet.get_all_values() 
df = pd.DataFrame(data) 

font= cv2.FONT_ITALIC
font1= cv2.FONT_HERSHEY_SIMPLEX
font = cv2.FONT_HERSHEY_COMPLEX

for i in range(2,len(df)):
  # Put the path of the certificate template in the google drive
  im=cv2.imread('/content/drive/MyDrive/JUIT/ECE/ISPCC2021/certificates/Presenter.jpg');

  # index 2 that is 3rd Column contains Name of the Participant
  txt1=df[2][i];
  # st_name points the starting pixel point from where name should start printing on the certificate
  st_name= round(1225-(len(txt1))/2*35);
  # 710 below indicates the row number of an image on which you want name of the participant to print
  cv2.putText(im,txt1,(st_name,710), font, 2,(200,115,50),3,cv2.LINE_AA);
  
  # index 3 that is 4th Column contains Affiliation of the Participant
  txt1=df[3][i];
  # st_name points the starting pixel point from where name should start printing on the certificate
  st_name= round(1225-(len(txt1))/2*35);
  # 710 below indicates the row number of an image on which you want name of the participant to print
  cv2.putText(im,txt1,(st_name,1010), font, 2,(200,115,50),3,cv2.LINE_AA);
  
  # index 0 that is 1st Column contains certificate number of the Participant
  txt3=df[0][i];
  cv2.putText(im,txt3,(2050,65), font1, 1,(0,0,0),2,cv2.LINE_AA);

  # It Displays the certificate below to check if the name and the certificate is printed properly or not
  plt.imshow(im)

  # To save the certicate generated above on the google drive at the path mentioned below. FIle name is saved with the name of the participant.
  pat='/content/drive/MyDrive/JUIT/ECE/ISPCC2021/certificates/Certificates/'+ df[2][i] + '.jpg';
  cv2.imwrite(pat,im)

# To check if everything is printing properly on the certificate, you can change the code below as comment so that email may not be send to the participants 

# index 4 that is 5th Column contains email id of the Participant
  receiver_email = df[4][i];
#   # Create a multipart message and set headers
  message = MIMEMultipart()
  message["From"] = sender_email
  message["To"] = receiver_email
  message["Subject"] = subject
  # message["Bcc"] = receiver_email  # Recommended for mass emails
  message.attach(MIMEText(body, "plain"))

  # Open image file in binary mode
  with open(pat, "rb") as attachment:
      # Add file as application/octet-stream
      # Email client can usually download this automatically as attachment
      part = MIMEBase("application", "octet-stream")
      part.set_payload(attachment.read())

  # Encode file in ASCII characters to send by email    
  encoders.encode_base64(part)

  # Add header as key/value pair to attachment part
  part.add_header(
      "Content-Disposition",
      f"attachment; filename= {filename1}",
  )

  # # Add attachment to message and convert message to string
  message.attach(part)
  text = message.as_string()

  # Log in to server using secure context and send email
  context = ssl.create_default_context()
  with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
      server.login(sender_email, password)
      server.sendmail(sender_email, receiver_email, text)

  del message
