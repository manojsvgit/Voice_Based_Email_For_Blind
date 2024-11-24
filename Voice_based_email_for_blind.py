import email
import imaplib
import os
import smtplib
import time
import pyglet
import speech_recognition as sr
from gtts import gTTS
from email.header import decode_header

# Email credentials (Set your email and password here)
SENDER_EMAIL = "vsjonam@gmail.com"
SENDER_PASSWORD = "tnugctbnurqwseco"


# Terminal visual enhancements
def styled_print(text, style="info"):
    styles = {
        "info": "\033[94m[INFO]\033[0m ",
        "success": "\033[92m[SUCCESS]\033[0m ",
        "error": "\033[91m[ERROR]\033[0m ",
        "prompt": "\033[93m[PROMPT]\033[0m ",
    }
    print(styles.get(style, "") + text)


# Function to speak the text
def speak(text):
    tts = gTTS(text=text, lang='en')
    ttsname = "output.mp3"
    tts.save(ttsname)
    music = pyglet.media.load(ttsname, streaming=False)
    music.play()
    time.sleep(music.duration)
    os.remove(ttsname)


# Recognize speech function
def recognize_speech():
    r = sr.Recognizer()
    while True:
        with sr.Microphone() as source:
            styled_print("Listening for your input...", "prompt")
            speak("Speak now.")
            audio = r.listen(source)
        try:
            text = r.recognize_google(audio)
            speak("You said: " + text)
            styled_print(f"You said: {text}", "success")
            return text.lower()
        except sr.UnknownValueError:
            speak("Sorry, I couldn't understand that. Would you mind repeating?")
            styled_print("Couldn't understand. Please try again.", "error")
        except sr.RequestError as e:
            speak(f"Error with the speech recognition service: {e}")
            styled_print(f"Speech recognition error: {e}", "error")


# Get recipient email address based on name
def get_recipient_email():
    speak("Please tell me the recipient's name.")
    recipient_name = recognize_speech()

    # Convert recipient's name to Gmail format (lowercase and space to no space)
    recipient_email = recipient_name.replace(" ", "").lower() + "@gmail.com"
    speak(
        f"The recipient's email address will be {recipient_email}. Is that correct? Say 'yes' to confirm or 'no' to change.")
    styled_print(
        f"The recipient's email address will be {recipient_email}. Is that correct? Say 'yes' to confirm or 'no' to change.",
        "prompt")

    confirmation = recognize_speech()
    if "yes" in confirmation:
        return recipient_email
    else:
        speak("Okay, please say the recipient's email address again.")
        styled_print("Please say the recipient's email address again.", "prompt")
        return get_recipient_email()


# Send email function (Your working method)
def send_email(sender_email, sender_password, recipient_email, message):
    try:
        styled_print("Connecting to the email server...", "info")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        styled_print("Logged in to email server.", "success")

        # Constructing the email
        email_message = f"Subject: Voice Email System\n\n{message}"
        server.sendmail(sender_email, recipient_email, email_message)
        speak("Your email has been sent successfully.")
        styled_print("Email sent successfully!", "success")
        server.quit()
    except Exception as e:
        speak("Failed to send the email. Please try again later.")
        styled_print(f"Failed to send email: {e}", "error")


# Check inbox function
def check_inbox(username, password):
    try:
        styled_print("Connecting to the email server to fetch inbox...", "info")
        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        mail.login(username, password)
        mail.select('inbox')

        # Fetch total and unread email count
        status, messages = mail.search(None, 'ALL')
        total_mails = len(messages[0].split())
        speak(f"There are {total_mails} emails in your inbox.")
        styled_print(f"Total emails: {total_mails}", "info")

        status, unseen_messages = mail.search(None, 'UNSEEN')
        unseen_count = len(unseen_messages[0].split())
        speak(f"There are {unseen_count} unread emails.")
        styled_print(f"Unread emails: {unseen_count}", "info")

        # Read the latest unread email
        if unseen_count > 0:
            result, data = mail.fetch(unseen_messages[0].split()[0], '(RFC822)')
            raw_email = data[0][1].decode("utf-8")
            email_message = email.message_from_string(raw_email)

            # Extract subject and sender
            subject, encoding = decode_header(email_message["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")
            from_ = email_message["From"]

            speak(f"Email from: {from_}")
            styled_print(f"From: {from_}", "info")
            speak(f"Subject: {subject}")
            styled_print(f"Subject: {subject}", "info")

            # Extract body
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        body = part.get_payload(decode=True).decode()
                        break
            else:
                body = email_message.get_payload(decode=True).decode()

            speak(f"Body: {body}")
            styled_print(f"Body: {body}", "info")
        else:
            speak("No unread emails.")
            styled_print("No unread emails.", "info")

        mail.close()
        mail.logout()
    except Exception as e:
        speak("Error checking inbox. Please try again later.")
        styled_print(f"Error checking inbox: {e}", "error")


# Confirm email address
def confirm_email_address(recipient_email):
    speak(f"Are you sure you want to send the email to {recipient_email}? Say 'yes' to confirm or 'no' to change.")
    styled_print(
        f"Are you sure you want to send the email to {recipient_email}? Say 'yes' to confirm or 'no' to change.",
        "prompt")
    confirmation = recognize_speech()

    if "yes" in confirmation:
        return True
    else:
        speak("Okay, please provide a different email address.")
        styled_print("Please provide a different email address.", "prompt")
        return False


# Main function
def main():
    # Greeting and user authentication
    styled_print("Voice-Based Email for Blind", "info")
    sender_email = "vsjonam@gmail.com"
    speak(f"Hello! Your email address is {sender_email}. Welcome to Voice-based Email for Blind.")

    while True:
        # Display options
        styled_print("What would you like to do?", "prompt")
        styled_print("1. Compose an email.", "prompt")
        styled_print("2. Check your inbox.", "prompt")
        styled_print("3. Exit.", "prompt")
        speak("Please choose an option. Compose, Inbox, or Exit.")
        choice = recognize_speech()

        if "compose" in choice or "email" in choice:
            # Compose email
            recipient_email = get_recipient_email()

            # Check if the recipient email is a valid Gmail address
            if "@gmail.com" not in recipient_email:
                speak("Sorry, we can only send emails to Gmail addresses.")
                styled_print("Invalid email. We can only send to Gmail addresses.", "error")
                continue

            # Confirm the email address
            if confirm_email_address(recipient_email):
                speak("Please say your email message.")
                message = recognize_speech()
                send_email(sender_email, SENDER_PASSWORD, recipient_email, message)

        elif "inbox" in choice:
            # Check inbox
            check_inbox(sender_email, SENDER_PASSWORD)

        elif "exit" in choice:
            speak("Thank you for using Voice-based Email for Blind. Have a Great Day!")
            styled_print("Exiting the application. Have a great day!", "success")
            break

        else:
            speak("Invalid choice. Please try again.")
            styled_print("Invalid choice. Please try again.", "error")


if __name__ == "__main__":
    main()
