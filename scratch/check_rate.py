import pyttsx3
engine = pyttsx3.init()
rate = engine.getProperty('rate')
print(f"Default rate: {rate}")
voices = engine.getProperty('voices')
for v in voices:
    print(f"Voice: {v.name}, ID: {v.id}")
