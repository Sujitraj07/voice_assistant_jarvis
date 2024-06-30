import requests
import win32com.client
import speech_recognition as sr
import datetime
import cv2
import threading
import webbrowser

speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Either provide the OpenWeatherAPI key directly to the API_KEY
API_KEY = ""


def say(text):
    s = text
    speaker.Speak(s)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry Sir..."


# Function to open and manage camera feed
def open_camera():
    global camera_open, cap, recording, output, frame
    try:
        say("Opening the camera...")
        cap = cv2.VideoCapture(0)
        camera_open = True
        recording = False
        output = None
        frame = None
        while camera_open:
            ret, frame = cap.read()
            if ret:
                if recording:
                    if output is None:
                        fourcc = cv2.VideoWriter_fourcc(*'XVID')
                        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        video_filename = f"recorded_{timestamp}.avi"
                        output = cv2.VideoWriter(video_filename, fourcc, 20.0, (640, 480))
                    output.write(frame)

                cv2.imshow("Camera Feed", frame)
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
    except Exception as e:
        print(f"An error occurred while opening the camera: {e}")
    finally:
        if cap is not None:
            cap.release()
            cv2.destroyAllWindows()
            if output is not None:
                output.release()
            camera_open = False


def open_website(website_name):
    websites = {
        "youtube": "https://www.youtube.com",
        "wikipedia": "https://www.wikipedia.org",
        "google": "https://www.google.com",
        "amazon": "https://www.amazon.com",
        "chaibasa engineering college": "https://chaibasaengg.edu.in"
    }

    if website_name.lower() in websites:
        url = websites[website_name.lower()]
        webbrowser.open(url)
        say(f"Opening {website_name} in your web browser.")
    else:
        say(f"Sorry, I don't have information on how to open {website_name}.")


def get_weather(city_name):
    try:
        base_url = f"http://api.openweathermap.org/data/2.5/weather?q={city_name}&appid={API_KEY}&units=metric"
        response = requests.get(base_url)
        data = response.json()

        if data["cod"] == 200:
            weather_data = data["weather"][0]
            main_data = data["main"]
            temperature = main_data["temp"]
            humidity = main_data["humidity"]
            description = weather_data["description"]

            weather_info = f"Weather in {city_name}: {description}, Temperature: {temperature}°C, Humidity: {humidity}%"
            return weather_info
        else:
            return "City not found"

    except Exception as e:
        return f"An error occurred while fetching weather information: {e}"


if __name__ == '__main__':
    say("Hello, I am Jarvis AI")
    camera_open = False
    cap = None
    camera_thread = threading.Thread(target=open_camera)

    while True:
        print("Listening...")
        text = takeCommand()
        if "open camera" in text and not camera_open:
            if camera_thread.is_alive():
                say("Camera is already open.")
            else:
                say("Opening the camera...")
                camera_thread.start()
        elif "close camera" in text and camera_open:
            say("Closing the camera...")
            camera_open = False
            if recording:
                say("Stopping video recording...")
                recording = False
                output.release()
            if camera_thread.is_alive():
                camera_thread.join()
        elif "capture" in text and camera_open:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            photo_filename = f"captured_{timestamp}.jpg"
            cv2.imwrite(photo_filename, frame)
            say(f"Picture captured as {photo_filename}")
        elif "start recording" in text and camera_open:
            if not recording:
                say("Starting video recording...")
                recording = True
                fourcc = cv2.VideoWriter_fourcc(*'XVID')
                audio_fourcc = cv2.VideoWriter_fourcc(*'MPEG')
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                video_filename = f"recorded_{timestamp}.avi"
                output = cv2.VideoWriter(video_filename, fourcc, 20.0, (640, 480), isColor=True)
        elif "stop recording" in text and camera_open:
            if recording:
                say("Stopping video recording...")
                recording = False
                if output is not None:
                    output.release()

        elif "search the web" in text:
            parts = text.split("for", 1)
            query = parts[1].strip()
            webbrowser.open(f"https://www.google.com/search?q={query}")


        elif "open website" in text:
            words = text.split()
            website_name = words[-1]
            open_website(website_name)

        elif "weather updates" in text:
            try:
                say("Sure, please tell me the city for weather updates.")
                city = takeCommand()
                weather_info = get_weather(city)
                say(weather_info)
            except Exception as e:
                say(f"An error occurred while fetching weather information: {e}")
        else:
            say(text)