import win32com.client
import speech_recognition as sr
import webbrowser as wb
import openai
import datetime
import os
import random
import csv
import codecs
import urllib.request
import urllib.error
import sys

from config import apikey


import requests
url = ('https://newsapi.org/v2/top-headlines?'
       'country=in&'
       'apiKey=54e9b6501b464fc9b583a604496f1897')
response = requests.get(url)



def weatherTell(location):

    # This is the core of our weather query URL
    BaseURL = 'https://weather.visualcrossing.com/VisualCrossingWebServices/rest/services/timeline/'

    ApiKey = 'G7PV68QCFRS3RLHEVGX7C4UQP'
    # UnitGroup sets the units of the output - us or metric
    UnitGroup = 'us'

    # Location for the weather data
    Location = f'{location}'

    # Optional start and end dates
    # If nothing is specified, the forecast is retrieved.
    # If start date only is specified, a single historical or forecast day will be retrieved
    # If both start and and end date are specified, a date range will be retrieved
    StartDate: str = ''
    EndDate = ''

    # JSON or CSV
    # JSON format supports daily, hourly, current conditions, weather alerts and events in a single JSON package
    # CSV format requires an 'include' parameter below to indicate which table section is required
    ContentType = "csv"

    # include sections
    # values include days,hours,current,alerts
    Include = "days"

    # basic query including location
    ApiQuery = BaseURL + Location

    # append the start and end date if present
    if len(StartDate):
        ApiQuery += "/" + StartDate
        if len(EndDate):
            ApiQuery += "/" + EndDate

    # Url is completed. Now add query parameters (could be passed as GET or POST)
    ApiQuery += "?"

    # append each parameter as necessary
    if len(UnitGroup):
        ApiQuery += "&unitGroup=" + UnitGroup

    if len(ContentType):
        ApiQuery += "&contentType=" + ContentType

    if len(Include):
        ApiQuery += "&include=" + Include

    ApiQuery += "&key=" + ApiKey

    print(' - Running query URL: ', ApiQuery)
    print()

    try:
        CSVBytes = urllib.request.urlopen(ApiQuery)
    except urllib.error.HTTPError as e:
        ErrorInfo = e.read().decode()
        print('Error code: ', e.code, ErrorInfo)
        sys.exit()
    except urllib.error.URLError as e:
        ErrorInfo = e.read().decode()
        print('Error code: ', e.code, ErrorInfo)
        sys.exit()

    # Parse the results as CSV
    CSVText = csv.reader(codecs.iterdecode(CSVBytes, 'utf-8'))

    RowIndex = 0

    # The first row contain the headers and the additional rows each contain the weather metrics for a single day
    # To simply our code, we use the knowledge that column 0 contains the location and column 1 contains the date.  The data starts at column 4
    for Row in CSVText:
        if RowIndex == 0:
            FirstRow = Row
        else:
            print('Weather in ', Row[0], ' on ', Row[1])

            ColIndex = 0
            for Col in Row:
                if ColIndex >= 4:
                    print('   ', FirstRow[ColIndex], ' = ', Row[ColIndex])
                ColIndex += 1
        RowIndex += 1

    # If there are no CSV rows then something fundamental went wrong
    if RowIndex == 0:
        print('Sorry, but it appears that there was an error connecting to the weather server.')
        print('Please check your network connection and try again..')

    # If there is only one CSV  row then we likely got an error from the server
    if RowIndex == 1:
        print('Sorry, but it appears that there was an error retrieving the weather data.')
        print('Error: ', FirstRow)

    return()


speaker = win32com.client.Dispatch("Sapi.SpVoice")

chatStr = ""


# https://youtu.be/Z3ZAJoi4x6Q
def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"{name}: {query}\n Iris: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap this inside of a  try catch block
    say(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]


def say(i):
    speaker.Speak(i)


def ai(prompt):
    openai.api_key = apikey
    text = f"Openai response for prompt : {prompt} \n ************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/prompt -  {random.randint(123413444, 19349963357)}.txt", "w") as f:
        f.write(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occurred. Sorry from Iris"


if __name__ == '__main__':
    say("Hello I am Iris A.I. . What is your name ?")
    print("Listening...")
    name = takeCommand()
    say(f"Hi {name}. what can i do for you")
    print(f"Hi {name}")
    while True:
        print("Listening...")
        query = takeCommand()
        sites = [["youtube", "https://youtube.com"],
                 ["facebook", "https://facebook.com"],
                 ["whatsapp", "https://web.whatsapp.com/"],
                 ["google", "https://google.com"],
                 ["wikipedia", "https://wikipedia.com"],
                 ["chat GPT", "https://chat.openai.com/"],
                 ["telegram", "https://t.me/"],
                 ["twitter", "https://twitter.com/"],
                 ["instagram", "https://www.instagram.com/"],
                 ["github", "https://github.com/"],
                 ['linkedin', "https://www.linkedin.com/"],
                 ['stackoverflow', "https://stackoverflow.com/"],
                 ["reddit", "https://www.reddit.com/"],
                 ['pinterest', "https://www.pinterest.com/"],
                 ['stackoverflow', "https://stackoverflow.com/"],
                 ['discord', 'https://discord.com/'],
                 ['spotify', 'https://open.spotify.com/']
                ]
        for site in sites:
            if f"open {site[0]}" in query.lower():
                say(f"Opening {site[0]} sir ...")
                wb.open(site[1])
            elif "tell the time" in query:
                strfTime = datetime.datetime.now().strftime("%H:%M")
                say(f"Sir the time is {strfTime}")
            elif "Using ai".lower() in query.lower():
                ai(prompt=query)
            elif "What is the temperature in".lower() in query.lower():
                say("Temperature in which city ?")
                location = takeCommand()
                wea = weatherTell(location)
                say(f"Weather in {location} is {wea}")
            elif "Tell me the news".lower() in query.lower():
                print("Narrating today's trending news")
                say(response.json())
            elif "Iris Quit".lower() in query.lower():
                exit()

            elif "reset chat".lower() in query.lower():
                chatStr = ""

            else:
                print("Chatting...")
                chat(query)
