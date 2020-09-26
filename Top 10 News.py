def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    import requests
    import json

    restart = ('y')
    print("Welcome to News Top 10")
    speak("Welcome to News Top 10")
    while restart == 'y':

        print("\n Press 1 for Business News \n Press 2 for Entertainment News \n "
              "Press 3 for Sports News \n Press 4 for Health News \n "
              "Press 5 for Science News \n Press 6 for Technology News \n")

        speak("Press 1 for Business News \n Press 2 for Entertainment News \n "
              "Press 3 for Sports News \n Press 4 for Health News \n "
              "Press 5 for Science News \n Press 6 for Technology News")

        speak("Enter Your Choice")
        n = int(input("Enter your choice : "))

        if n == 1:
            speak("Top 10 Headlines of Business News are ")
            url = ('http://newsapi.org/v2/top-headlines?country=in&category=business&apiKey=4e7cc567511b4388b652117a2c0cb5e7')
            news = requests.get(url).text
            my_json = json.loads(news)
            for i in range(0, 10):
                speak(my_json['articles'][i]['title'])
                if i < 9:
                    speak("Next News is...")
                else:
                    speak("Thank You For listening")
            speak("Press y to restart and any key to exit")
            restart = input("Enter y to restart : ").lower()
            if restart != 'y':
                speak("stay tuned with News Top 10")

        elif n == 2:
            speak("Top 10 Headlines of Entertainment News are ")
            url = ('http://newsapi.org/v2/top-headlines?country=in&category=entertainment&apiKey=4e7cc567511b4388b652117a2c0cb5e7')
            news = requests.get(url).text
            my_json = json.loads(news)
            for i in range(0, 10):
                speak(my_json['articles'][i]['title'])
                if i < 9:
                    speak("Next News is...")
                else:
                    speak("Thank You For listening")
            speak("Press y to restart and any key to exit")
            restart = input("Enter y to restart : ").lower()
            if restart != 'y':
                speak("stay tuned with News Top 10")

        elif n == 3:
            speak("Top 10 Headlines of Sports News are ")
            url = ('http://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=4e7cc567511b4388b652117a2c0cb5e7')
            news = requests.get(url).text
            my_json = json.loads(news)
            for i in range(0, 10):
                speak(my_json['articles'][i]['title'])
                if i < 9:
                    speak("Next News is...")
                else:
                    speak("Thank You For listening")
            speak("Press y to restart and any key to exit")
            restart = input("Enter y to restart : ").lower()
            if restart != 'y':
                speak("stay tuned with News Top 10")

        elif n == 4:
            speak("Top 10 Headlines of Health News are ")
            url = ('http://newsapi.org/v2/top-headlines?country=in&category=health&apiKey=4e7cc567511b4388b652117a2c0cb5e7')
            news = requests.get(url).text
            my_json = json.loads(news)
            for i in range(0, 10):
                speak(my_json['articles'][i]['title'])
                if i < 9:
                    speak("Next News is...")
                else:
                    speak("Thank You For listening")
            speak("Press y to restart and any key to exit")
            restart = input("Enter y to restart : ").lower()
            if restart != 'y':
                speak("stay tuned with News Top 10")

        elif n == 5:
            speak("Top 10 Headlines of Science News are ")
            url = ('http://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=4e7cc567511b4388b652117a2c0cb5e7')
            news = requests.get(url).text
            my_json = json.loads(news)
            for i in range(0, 10):
                speak(my_json['articles'][i]['title'])
                if i < 9:
                    speak("Next News is...")
                else:
                    speak("Thank You For listening")
            speak("Press y to restart and any key to exit")
            restart = input("Enter y to restart : ").lower()
            if restart != 'y':
                speak("stay tuned with News Top 10")

        elif n == 6:
            speak("Top 10 Headlines of Technology News are ")
            url = ('http://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=4e7cc567511b4388b652117a2c0cb5e7')
            news = requests.get(url).text
            my_json = json.loads(news)
            for i in range(0, 10):
                speak(my_json['articles'][i]['title'])
                if i < 9:
                    speak("Next News is...")
                else:
                    print("Thank You For listening")
                    speak("Thank You For listening")
            print("Press y to restart and any key to exit")
            speak("Press y to restart and any key to exit")
            restart = input("Enter y to restart : ").lower()
            if restart != 'y':
                speak("stay tuned with News Top 10")
        else:
            print("Invalid Option")
            speak("Invalid Option")
            restart = ('y')