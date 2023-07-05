import requests
import json
import win32com.client as wincom

speak = wincom.Dispatch("SAPI.SpVoice")

user = f"Hi user , Welcome to Weather program , TO get Weather information of Your area, \n Kindly Enter City Name :"
speak.Speak(user)

while True:
    city=input("City Name : ")
    if (city=="done"):
                break
    
    url =f"http://api.weatherapi.com/v1/current.json?key=b369c94e6e6a4bad9d4105004230507&q={city}"
    r = requests.get(url)

    wdic =json.loads(r.text)        # convert Json to string
    c = wdic["current"]["temp_c"]
    f= wdic["current"]["temp_f"]
    text= wdic["current"]["condition"]["text"]

    report = f"temperature of {city} is {c} degree celcius and it seems to be {text} in the region."
    speak.Speak( report)
    print('city :',city,"\nTemperature : ",c,'\ncondition :',text)
    print('Thank you !!  Kindly enter next city name or type done')

