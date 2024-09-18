import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
from dotenv import load_dotenv, find_dotenv
from datetime import datetime, timedelta
import pytz

# Load environment variables
keys = ['email_user', 'email_password', 'email_to', 'email_to_2']

for key in keys:
    if key in os.environ:
        del os.environ[key]

try:
    load_dotenv(find_dotenv())
    print("Environment variables loaded successfully")
except Exception as e:
    print(f"An error occurred while loading the environment variables: {e}")

# Environment variables for email credentials
email_user = os.getenv("email_user")
email_password = os.getenv("email_password")
email_to = os.getenv("email_to")
email_to_2 = os.getenv("email_to_2")

print(f"email_user is {email_user}")
print(f"email_to is {email_to}")
print(f"email_to_2 is {email_to_2}")

# Define time zone for Ecuador/Guayaquil
ECUADOR_TZ = pytz.timezone('America/Guayaquil')

def filter_recent_entries(data, current_time):
    # Define the cutoff time as 4 PM of the previous day in Ecuador time zone
    cutoff_time = current_time.replace(hour=16, minute=0, second=0, microsecond=0) - timedelta(days=1)
    
    filtered_data = []
    for entry in data:
        # Ensure the 'created_at' field is a string before parsing
        created_at = entry['created_at']
        if isinstance(created_at, pd.Timestamp):
            created_at = created_at.strftime('%Y-%m-%d %H:%M:%S')
        elif not isinstance(created_at, str):
            print(f"Skipping entry with invalid 'created_at': {created_at}")
            continue

        # Parse the datetime string and compare it to the cutoff
        entry_datetime = datetime.strptime(created_at, '%Y-%m-%d %H:%M:%S').replace(tzinfo=ECUADOR_TZ)
        if entry_datetime > cutoff_time:
            filtered_data.append(entry)

    return filtered_data

def json_to_csv(json_path, csv_path):
    try:
        # Read the JSON file
        with open(json_path, 'r') as f:
            data = pd.read_json(f)
        
        # List of keys of interest
        keys_of_interest = ['url', 'text', 'comment', 'user', 'user_profile', 'created_at', 'clasificacion', 'red_social', 'comment_id', 'tweet_id']
        
        # Filter entries based on recent entries criteria
        filtered_entries = filter_recent_entries(data.to_dict(orient='records'), datetime.now(ECUADOR_TZ))
        
        # Check which keys are present in the filtered data
        if filtered_entries:
            present_keys = [key for key in keys_of_interest if key in filtered_entries[0]]
            
            # Create DataFrame using only the present keys
            df = pd.DataFrame(filtered_entries)[present_keys]
            
            # Save to CSV
            df.to_csv(csv_path, index=False)
            print(f"Converted {json_path} to {csv_path} with recent entries and available keys.")
        else:
            print("No recent entries found; CSV not updated.")
    except Exception as e:
        print(f"Failed to convert JSON to CSV: {e}")


def send_email_with_attachment(attachment_path):
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = f"{email_to}, {email_to_2}"
    msg['Subject'] = f"Consolidado de Comentarios Banco Guayaquil {datetime.now(ECUADOR_TZ).strftime('%Y-%m-%d')}"

    TODAY = datetime.now(ECUADOR_TZ).strftime('%Y-%m-%d')

    body = f"Consolidado de los comentarios de todas las redes sociales del {TODAY}\nEl corte va desde las 4pm del dia anterior hasta las 4pm del dia de hoy.\n\nSaludos,\nDinamicDataLab"

    # Attach the body of the email
    msg.attach(MIMEText(body, 'plain'))

    # Attempt to attach the CSV file
    try:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            msg.attach(part)
    except FileNotFoundError:
        print(f"File not found: {attachment_path}. No daily comments available.")
        return  # Skip sending the email if the file is missing

    # Send the email
    try:
        print("Attempting to send email...")
        with smtplib.SMTP('smtp.office365.com', 587) as server:  # Ensure Outlook SMTP server
            server.starttls()
            server.login(email_user, email_password)
            server.send_message(msg)
            print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

def reset_csv(csv_path):
    try:
        # Empty the CSV by writing just the headers
        with open(csv_path, 'w') as f:
            f.write('')
        print("CSV file reset successfully.")
    except Exception as e:
        print(f"Failed to reset CSV file: {e}")

if __name__ == "__main__":
    # Define paths
    json_file_path = 'all_comments.json'
    csv_file_path = 'all_comments.csv'

    # Print current time for debugging
    current_time = datetime.now(ECUADOR_TZ)
    print(f"Current time: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")

    # Convert JSON to CSV with the recent entries
    json_to_csv(json_file_path, csv_file_path)

    # Send email with the CSV attachment
    send_email_with_attachment(csv_file_path)

    # Check current time to reset CSV at 4:00 PM
    if current_time.hour == 16 and current_time.minute == 0:  # Run reset at exactly 4:00 PM
        reset_csv(csv_file_path)
