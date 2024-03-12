from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from bs4 import BeautifulSoup
import requests
import base64
import os
from PIL import Image
import pytesseract


# Function to process and download images, then extract codes
def process_emails(query):
    results = service.users().messages().list(userId="me", q=query).execute()
    messages = results.get("messages", [])

    if messages:
        for i, message in enumerate(messages):
            message_id = message["id"]
            print(f"Processing message ID: {message_id}")  # Debugging output
            try:
                msg = (
                    service.users()
                    .messages()
                    .get(userId="me", id=message_id, format="full")
                    .execute()
                )

                # Extract email subject for filename
                headers = msg["payload"]["headers"]
                subject = next(
                    (
                        header["value"]
                        for header in headers
                        if header["name"].lower() == "subject"
                    ),
                    "NoSubject",
                )
                subject_sanitized = subject.split()[
                    0
                ]  # Use the first word of the subject
                subject_sanitized = "".join(
                    filter(str.isalnum, subject_sanitized)
                )  # Remove non-alphanumeric characters

                parts = msg.get("payload", {}).get("parts", [])

                html_part = None
                for part in parts:
                    if (
                        part.get("mimeType") == "text/html"
                        and "body" in part
                        and "data" in part["body"]
                    ):
                        html_part = base64.urlsafe_b64decode(
                            part["body"]["data"].encode("utf-8")
                        ).decode("utf-8")
                        break

                if html_part:
                    soup = BeautifulSoup(html_part, "html.parser")
                    images = soup.find_all("img")
                    for j, img in enumerate(images):
                        img_url = img["src"]
                        response = requests.get(img_url)
                        if response.status_code == 200:
                            # Construct filename using the subject and index
                            filename = f"{subject_sanitized}_{i}_{j}.png"
                            filepath = os.path.join("images", filename)
                            with open(filepath, "wb") as f:
                                f.write(response.content)
                            print(f"Downloaded {filename}")

            except Exception as e:
                print(f"An error occurred: {e}")
    else:
        print("No emails found.")


# Functions to extract codes from downloaded images
def extract_text_from_image(image_path):
    try:
        image = Image.open(image_path)
        text = pytesseract.image_to_string(image)
        processed_text = text.split()
        return processed_text
    except Exception as e:
        print(f"Error processing image: {e}")
        return []


def get_FIT3161_codes():
    processed_text = extract_text_from_image("images/FIT3161_0_0.png")
    if processed_text:
        seminar_code = processed_text[6]
        tutorial_code = processed_text[-1]
        print("\nFIT3161 Attendance codes")
        print(f"Seminar code: {seminar_code}\nTutorial code: {tutorial_code}\n")


def get_FIT3171_codes():
    processed_text_workshop = extract_text_from_image("images/FIT3171_0_0.png")
    processed_text_applied = extract_text_from_image("images/FIT3171_0_1.png")
    if processed_text_applied and processed_text_workshop:
        applied_code = processed_text_applied[-6]
        workshop_code = processed_text_workshop[6]

        print("\nFIT3171 Attendance codes")
        print(f"Workshop code: {workshop_code}\nApplied code: {applied_code}\n")


def get_ETW2001_codes():
    processed_text = extract_text_from_image("images/ETW2001_0_0.png")
    if processed_text:
        tutorial_code = processed_text[-5]
        print("\nETW2001 Attendance codes")
        print(f"Tutorial code: {tutorial_code}\n")


# Main execution starts here
# Load credentials and build the Gmail service
creds = Credentials.from_authorized_user_file("gmail_token.json")
service = build("gmail", "v1", credentials=creds)

# Define your search queries and corresponding processing functions
queries = [
    'subject:"FIT3161 - FIT3162 - MUM S1 2024 - Announcements - [ATTENDANCE] Week 3 codes"',
    'subject:"FIT3171 DATABASES - MUM S1 2024 - Announcements - [Week 3] Attendance Codes for International Students"',
    'subject:"ETW2001 - MUM S1 2024 - Announcements - Week 3 Attendance Codes (for international students)"',
]

functions = [get_FIT3161_codes, get_FIT3171_codes, get_ETW2001_codes]

# Process each query
for query in queries:
    process_emails(query)

for func in functions:
    func()
