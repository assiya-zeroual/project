import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import pdfkit
from datetime import date
from pypdf import PdfReader, PdfWriter
import os
from pdf2docx import Converter
import mimetypes
import os
config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
emplois=pd.read_excel("resultats_finaux (1).xlsx")
user =pd.read_excel("user_file.xlsx")
emplois["Domaine"] = emplois["Domaine"].astype(str).str.lower().str.strip()
user["Domaine"] = user["Domaine"].astype(str).str.lower().str.strip()
content4=f"""
            <html>
            <head>
            <meta charset="UTF-8">
            </head>
            <body>
            {date.today().isoformat()}"""
def fun(nom,email,filehtml):
      with open(filehtml,'r') as fils:
            fichier=fils.read()
      fichiermod=fichier.replace('[Votre Nom]', nom)
      fichiermod=fichiermod.replace('[Email]',email)
      with open(filehtml,'w') as h:
          h.write(fichiermod)
      pdfkit.from_file(filehtml,'cv.pdf',configuration=config)
      cv=Converter('cv.pdf')
      cv.convert('cv.docx')
      cv.close()
      return 'cv.docx'
categoris=user["Domaine"].unique()
print(categoris)
for i in categoris:
   list1 = [mot.strip() for mot in i.split("/")]
   pattern = "|".join(list1)
   colone1 = emplois[emplois["Domaine"].str.contains(pattern, case=False, na=False, regex=True)]
   print(colone1['Domaine'])
   colone2 = user[user["Domaine"].str.strip().str.lower() == i.strip().lower()]
   if colone1.empty:
      continue  
   for j in range(len(colone2)):
    msg = EmailMessage()
    mot_de_passe='dhda fukw scgt gsiu'
    msg['Subject'] = 'Offre d’emploi3'
    msg['From'] = 'chachhanae8@gmail.com'
    msg['To'] = colone2.iloc[j]['Email']
    content1=f"""{date.today().isoformat()}"""
    emails = ""
    for k in range(len(colone1)):
         emails+=f"""<a href="{colone1.iloc[k]['card-job-detail-href']}">{colone1.iloc[k]['card-job-detail-href']}</a> """
    chemin2 = None
    if i=="informatique / réseaux":
        chemin2=fun(colone2.iloc[j]['Prenom'],colone2.iloc[j]['Email'],'cv_info.html')
    html_content1 = f"""
            <html>
            <head>
            <meta charset="UTF-8">
            </head>
            <body>"""
    html_content2=f"""<p>Bonjour {colone2.iloc[j]['Prenom']} {colone2.iloc[j]['Nom']},</p>
              <p>Vous avez indiqué comme domaine : <strong>{i}</strong>.</p>
              <p>Nous avons trouvé des offres correspondants à votre profil :</p>
              <p>Voici toutes nos offres disponibles </p>
               {emails}
              <div style="margin-top: 20px;">
                Cordialement,<br>
                L'équipe d' INFITAH
              </div>
            </body>
            </html>
            """
    html_content=  html_content1+html_content2
    content=content4+html_content2
    content4+=html_content2
    msg.add_alternative(html_content, subtype='html')
    if chemin2 is not None:

      with open(chemin2, 'rb') as f:
         file_data = f.read()

      mime_type, _ = mimetypes.guess_type(chemin2)
      if mime_type is None:
          mime_type = 'application/octet-stream'

      maintype, subtype = mime_type.split('/', 1)

      msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=os.path.basename(chemin2))
    try:
      with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login('chachhanae8@gmail.com', mot_de_passe)
        smtp.send_message(msg)
        print(f"Email envoyé à : {colone2.iloc[j]['Email']}")
    except Exception as e:
       print(f"Erreur d'envoi à {colone2.iloc[j]['Email']} : {e}")
pdfkit.from_string(content4, "rapport.pdf",configuration=config)
rapport=PdfReader("rapport.pdf")
ecrire=PdfWriter()
for page in rapport.pages:
   ecrire.add_page(page)
if os.path.exists('rapport_pr.pdf'):
   rapport_pr=PdfReader('rapport_pr.pdf')
   for page2 in rapport_pr.pages:
      ecrire.add_page(page2)
with open('rapport_pr.pdf','wb') as pdf:
      ecrire.write(pdf)
os.remove("rapport.pdf")
    

  

           
                  






