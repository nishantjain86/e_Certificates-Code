# e_Certificates-Code
The Code is in Python will help you in sending the e-Certificates to the large of participants in a single click

# First step is to Create & use App Passwords. 
Follow the link below to create App Password:
https://support.google.com/mail/answer/185833?hl=en

Python code will read the details of the Participants that is Name , Email ID and Certificate Number from the Google Sheet. You can modify the code to read details from excel file also.

Code also saves all the certificates emailed on the google drive path mentioned in the code.

# GOOGLE SHEET TEMPLATE 
First Column --> Certificate Number

Third Column --> Participant Name to be printed on the certificate

Fourth Column --> Affiliation of the Participant Name to be printed on the certificate

Fifth Column --> Email Address of the Participant

# PYTHON CODE 
## In the Code do the following changes as per the requirement:

## Put Email Subject
subject = "Certificate for  ISPCC 2021" 

## Put Email Body
body = "Dear Participant,\n        Thank You for presenting a paper in ISPCC-2021 organized by Department of Electronics and Communication Engineering, Jaypee University of Information Technology (JUIT), Waknaghat, Solan, H.P., INDIA."
body=body +"\n\nRegards \nOrganizing Committee \nISPCC 2021 \nDepartment of ECE\nJUIT, Waknaghat, Solan, HP, India"

## Put Sender Email ID and Password ( You need to generate password through google account for third party app)
sender_email = "nishantjain86@gmail.com"
password = "************" 

## Put the text with which you want to send the file/Certificate.
filename1="ISPCC-2021,JUIT.jpg"

## Put the URL of the spread sheet 
wb = gc.open_by_url('https://docs.google.com/spreadsheets/d/1mRhotki-RiN_jQvhkhjghjD71NBgisvj-NkeD_u33xE-QIMaaU/edit#gid=0') 

## Put the name of the sheet that contains the List of certificates 
sheet = wb.worksheet('a')

## Put the path of the certificate template in the google drive
im=cv2.imread('/content/drive/MyDrive/JUIT/ECE/ISPCC2021/certi/Presenter.jpg');

## Below data need to be updated as per the location at which information need to be printed on the Certificate
### st_name points the starting pixel point from where name should start printing on the certificate
  st_name= round(1225-(len(txt1))/2*35);

### 710 below indicates the row number of an image on which you want name of the participant to print
  cv2.putText(im,txt1,(st_name,710), font, 2,(200,115,50),3,cv2.LINE_AA);

## To save the certicate generated above on the google drive at the path mentioned below. FIle name is saved with the name of the participant.
  pat='/content/drive/MyDrive/JUIT/ECE/ISPCC2021/certificates/Certificates/'+ df[2][i] + '.jpg';
  
 


