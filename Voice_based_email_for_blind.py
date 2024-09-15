import email
import imaplib
import os
import smtplib
import time
import pyglet
import speech_recognition as sr
from bs4 import BeautifulSoup
from gtts import gTTS


def speak(text):
    tts = gTTS(text=text, lang='en')
    ttsname = "output.mp3"
    tts.save(ttsname)
    music = pyglet.media.load(ttsname, streaming=False)
    music.play()
    time.sleep(music.duration)
    os.remove(ttsname)


def recognize_speech():
    r = sr.Recognizer()
    while True:
        with sr.Microphone() as source:
            speak("Speak now:")
            audio = r.listen(source)
            speak("OK, done")
        try:
            text = r.recognize_google(audio)
            speak("You said: " + text)
            return text.lower()
        except sr.UnknownValueError:
            speak("Google Speech Recognition could not understand audio. Please try again.")
        except sr.RequestError as e:
            speak("Could not request results from Google Speech Recognition service; {0}".format(e))


def get_user_name():
    speak("May I know your name?")
    name = recognize_speech()
    return name


def send_email(sender_email, sender_password, recipient_email, message):
    try:
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()
        mail.starttls()
        mail.login(sender_email, sender_password)
        mail.sendmail(sender_email, recipient_email, message)
        speak("Congratulations! Your mail has been sent.")
        mail.close()
    except Exception as e:
        speak("Any other help you need")


def check_inbox(username, password):
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        mail.login(username, password)
        mail.select('inbox')
        # Get total number of mails
        status, messages = mail.search(None, 'ALL')
        total_mails = len(messages[0].split())
        speak("Number of mails in your inbox: " + str(total_mails))
        # Get unseen mails
        status, unseen_messages = mail.search(None, 'UNSEEN')
        unseen_count = len(unseen_messages[0].split())
        speak("Number of unseen mails: " + str(unseen_count))
        # Get the latest email
        result, data = mail.fetch(messages[0], '(RFC822)')
        raw_email = data[0][1].decode("utf-8")
        email_message = email.message_from_string(raw_email)
        speak("From: " + email_message['From'])
        speak("Subject: " + str(email_message['Subject']))
        # Get the body of the email
        status, message_data = mail.fetch(messages[0], '(UID BODY[TEXT])')
        raw_body = message_data[0][1]
        soup = BeautifulSoup(raw_body, "html.parser")
        body_text = soup.get_text()
        speak("Body: " + body_text)
        mail.close()
        mail.logout()
    except Exception as e:
        speak("That's it for today.")


def main():
    # Greeting
    name = get_user_name()
    speak(f"Hello {name}! Welcome to Voice-based Email for Blind.")
    # Choices
    speak("What would you like to do?")
    speak("1. Compose a mail.")
    speak("2. Check your inbox")
    # Voice recognition part
    choice = recognize_speech()
    # Choices details
    if choice == '1' or 'compose' in choice:
        speak("Please say your message:")
        message = recognize_speech()
        # Email configuration
        sender_email = 'vsjonam@gmail.com'
        sender_password = 'tnugctbnurqwseco'
        recipient_email = 'manojpoojary9242@gmail.com'
        # Send email
        send_email(sender_email, sender_password, recipient_email, message)
    elif choice == '2' or 'inbox' in choice:
        # Email configuration
        username = 'vsjonam@gmail.com'
        password = 'tnugctbnurqwseco'
        # Check inbox
        check_inbox(username, password)
        speak("Thank you for using Voice-based Email for Blind. Have a great day!")


if __name__ == "__main__":
    main()
