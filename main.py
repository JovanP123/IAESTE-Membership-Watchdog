import pandas as pd
import json
import requests
from io import BytesIO

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

config_path = "config.json"
config_loaded_json = None

days_left_column_index = 4


succesful_response_status_code = 200
print("Reading config file...")
try:
    f = open(config_path, "r")
    config_loaded_json = json.load(f)
except FileNotFoundError:
    try:
        f = open("config.example.json", "r")
        config_loaded_json = json.load(f)
    except FileNotFoundError:
        print(f"Couldn't find {config_path} nor config.example.json...")
        exit()

def get_excel(url, username, password):
    session = requests.Session()
    session.auth = (username, password)
    response = session.get(url)

    if response.status_code == succesful_response_status_code:
        return BytesIO(response.content)
    else:
        raise Exception(f"Failed to fetch file: {response.status_code} - {response.text}")
def excel_to_array(excelFile):
    df = pd.read_excel(excelFile, engine="openpyxl")
    return df.values.tolist()
def send_mail_from_outlook(recipient_email, subject, body):
    sender_email = config_loaded_json["outlook_email"]
    app_password = config_loaded_json["outlook_password"]

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
            print(f"Email to {recipient_email} on subject {subject} has been sent successfully")
    except Exception as e:
        print(f"FAILED: {e}")

print("Getting excel file from sharepoint...")
excel_file = get_excel(config_loaded_json["sharepoint_excel_url"],config_loaded_json["sharepoint_username"],config_loaded_json["sharepoint_password"])
excel_file_array = excel_to_array(excel_file)

members_count = 0
at_member = 1
while excel_file_array[0][at_member] != "":
    at_member += 1
    members_count += 1

for i in range(1, members_count+1):
    if excel_file_array[days_left_column_index][i] == "7":
        #prati mejlcence

