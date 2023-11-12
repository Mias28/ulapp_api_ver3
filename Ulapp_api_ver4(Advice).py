import requests
import json
import pandas as pd
import os

# API of open weather app
API_KEY = 'c3ab9a2a4ccfaac2065c44ce3977630b'

def get_5_day_forecast(city_name):
    base_url = 'http://api.openweathermap.org/data/2.5/forecast'
    
    # Parameters that bring sorrow
    params = {
        'q': city_name,
        'appid': API_KEY,
        'units': 'metric',
    }
    
    # Request data (I might end up in the psych ward)
    response = requests.get(base_url, params=params)
    
    if response.status_code == 200:
        data = response.json()
        return data
    else:
        print(f"Error: {response.status_code}")
        return None

def display_5_day_forecast(data):
    if data:
        city = data['city']['name']
        forecast_list = data['list']

        # Keeps track of the days
        current_day = None
        day_counter = 0

        # Create an Excel writer with the specified filename
        save_directory = r"D:\Hist_data\\"
        excel_filename = f"{save_directory}{city}_5_day_forecast.xlsx"

        with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
            forecast_data = []  # Initialize the list here

            for item in forecast_list:
                timestamp = item['dt_txt']
                date = timestamp.split()[0]  # Extract date from the API

                # Checker for a new day in the forecast
                if date != current_day:
                    if current_day is not None:
                        # Save the DataFrame to a new sheet
                        df = pd.DataFrame(forecast_data)
                        df.to_excel(writer, index=False, sheet_name=f'Day {current_day}')

                    current_day = date
                    day_counter += 1
                    print(f"Day {day_counter}:")

                    # Empty list to store data for the current day
                    forecast_data = []

                day_data = {
                    'Date & Time': timestamp,
                    'Temperature (°C)': item['main']['temp'],
                    'Description': item['weather'][0]['description'],
                    'Humidity (%)': item['main']['humidity'],
                    'Wind Speed (m/s)': item['wind']['speed'],
                    'Rain (3h) (mm)': item.get('rain', {}).get('3h', 0)
                }
                forecast_data.append(day_data)

                temperature = item['main']['temp']
                description = item['weather'][0]['description']
                humidity = item['main']['humidity']
                wind_speed = item['wind']['speed']
                rain_3h = item.get('rain', {}).get('3h', 0)

                print(f"Date & Time: {timestamp}")
                print(f"Temperature: {temperature}°C")
                print(f"Description: {description}")
                print(f"Humidity: {humidity}%")
                print(f"Wind Speed: {wind_speed} m/s")
                print(f"Rain (3h): {rain_3h} mm")

                # Provide advice based on the weather description
                advice(description)

                print()  # Add an empty line for better readability

            # Save the last day's data to a new sheet
            df = pd.DataFrame(forecast_data)
            df.to_excel(writer, index=False, sheet_name=f'Day {current_day}')

        save_to_excel = input("Do you want to save the 5-Day Weather Forecast to Excel? (yes/no): ").lower()

        if save_to_excel == 'yes':
            print(f"5-Day Weather Forecast saved to {excel_filename}")
        else:
            os.remove(excel_filename)
            print("Forecast data not saved.")
    else:
        print("No data found :(")

def advice(weather_description):
    # Provides advice based on the weather description
    if 'clear sky' in weather_description.lower():
        print("Enjoy the clear skies!")
    elif 'few clouds' in weather_description.lower():
        print("It's a beautiful day with a few clouds.")
    elif 'scattered clouds' in weather_description.lower():
        print("Expect some scattered clouds in the sky.")
    elif 'broken clouds' in weather_description.lower():
        print("The sky has broken clouds today. Enjoy your day.")
    elif 'overcast clouds' in weather_description.lower():
        print("It's overcast today. Don't forget your umbrella.")
    elif 'light rain' in weather_description.lower():
        print("Light rain expected. Bring an umbrella.")
    elif 'moderate rain' in weather_description.lower():
        print("Moderate rain expected. Stay dry by bringing a umbrella or rain coat.")
    elif 'heavy rain' in weather_description.lower():
        print("Heavy rain expected. Bring an umbrella and avoid open areas.")
    elif 'drizzle' in weather_description.lower():
        print("Drizzle expected. A light raincoat might be handy.")
    elif 'thunderstorm' in weather_description.lower():
        print("Thunderstorm alert! Stay indoors.")
    elif 'snow' in weather_description.lower():
        print("Snowfall expected. Wear heavy clothing.")
    elif 'mist' in weather_description.lower():
        print("Misty weather. Drive safely.")
    elif 'fog' in weather_description.lower():
        print("Foggy conditions. Drive with caution.")
    elif 'smoke' in weather_description.lower():
        print("Smoky conditions. Stay indoors if possible.")
    elif 'haze' in weather_description.lower():
        print("Hazy weather. Limit outdoor activities.")
    elif 'dust' in weather_description.lower():
        print("Dusty conditions. Wear a face mask.")
    else:
        print("Check for specific weather updates.")

if __name__ == '__main__':
    print('Ulapp Forecasting(BETA)')
    city_name = input("Enter city name: ")
    forecast_data = get_5_day_forecast(city_name)
    display_5_day_forecast(forecast_data)
