import smtplib
import requests
import pywhatkit as pwk
import geocoder
import pyttsx3
from bs4 import BeautifulSoup
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Function to send an email
def send_email(subject, body, to_email):
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port =587
        from_email ="enter sender email"
        password ="enter yout password"
        message = f'Subject: {subject}\n\n{body}'

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, password)
            server.sendmail(from_email, to_email, message)
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Function to send an SMS
def send_sms(phone_number, message):
    try:
        pwk.sendwhatmsg_instantly(phone_number, message)
        print("SMS sent successfully")
    except Exception as e:
        print(f"Failed to send SMS: {e}")

# Function to scrape top 5 search results from Google
def google_search(query):
    try:
        url = f"https://www.google.com/search?q={query}"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')

        results = soup.find_all('h3', limit=5)
        for index, result in enumerate(results, 1):
            print(f"{index}. {result.text}")

    except Exception as e:
        print(f"Failed to scrape Google search results: {e}")

# Function to find current geo coordinates and location
def get_geo_coordinates():
    try:
        g = geocoder.ip('me')
        print(f"Your current location: {g.city}, {g.state}, {g.country}")
        print(f"Coordinates: {g.latlng}")
    except Exception as e:
        print(f"Failed to get geo coordinates: {e}")

# Function to convert text to audio
def text_to_audio(text):
    try:
        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"Failed to convert text to audio: {e}")

# Function to control volume of your laptop
def set_volume(level):
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        # Normalize the level (input 0-100) to the range that pycaw uses (-96.0 to 0.0 dB)
        level_normalized = level / 100.0
        volume.SetMasterVolumeLevelScalar(level_normalized, None)
        print(f"Volume set to {level}%")

    except Exception as e:
        print(f"Failed to control volume: {e}")


# Function to send bulk emails
def send_bulk_emails(subject, body, email_list):
    for email in email_list:
        send_email(subject, body, email)

# Main function to call all other functions
def main():
    print("Choose a task:")
    print("1. Send Email")
    print("2. Send SMS")
    print("3. Scrape Google Search Results")
    print("4. Get Geo Coordinates")
    print("5. Convert Text to Audio")
    print("6. Control Volume")
    print("7. Send Bulk Emails")

    choice = int(input("Enter your choice (1-7): "))

    if choice == 1:
        subject = input("Enter the subject: ")
        body = input("Enter the email body: ")
        to_email = input("Enter the recipient's email: ")
        send_email(subject, body, to_email)

    elif choice == 2:
        phone_number = input("Enter the phone number: ")
        message = input("Enter the message: ")
        send_sms(phone_number, message)

    elif choice == 3:
        query = input("Enter the search query: ")
        google_search(query)

    elif choice == 4:
        get_geo_coordinates()

    elif choice == 5:
        text = input("Enter the text to convert to audio: ")
        text_to_audio(text)

    elif choice == 6:
        level = int(input("Enter the volume level (0-100): "))
        set_volume(level)

    elif choice == 7:
        subject = input("Enter the subject: ")
        body = input("Enter the email body: ")
        email_list = input("Enter the recipient emails separated by comma: ").split(',')
        send_bulk_emails(subject, body, email_list)

    else:
        print("Invalid choice")

if __name__ == "__main__":
    main()