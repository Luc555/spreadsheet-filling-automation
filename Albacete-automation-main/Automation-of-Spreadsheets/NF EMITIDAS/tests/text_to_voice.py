"""
import pyttsx3  
s = pyttsx3.init()  
data = "Lucas, I love you so much"  
s.say(data)  
s.runAndWait()

from pygame import mixer
from gtts import gTTS
import time
import os

if os.name == "posix":
    from pydub import AudioSegment


# PART 1
def say(text, filename="temp.mp3", delete_audio_file=True, language="en", slow=False):
    # PART 2
    audio = gTTS(text, lang=language, slow=slow)
    audio.save(filename)

    if os.name == "posix":
        sound = AudioSegment.from_mp3(filename)
        old_filename = filename
        filename = filename.split(".")[0] + ".ogg"
        sound.export(filename, format="ogg")
        if delete_audio_file:
            os.remove(old_filename)

    # PART 3
    mixer.init()
    mixer.music.load(filename)
    mixer.music.play()

    # PART 4
    seconds = 0
    while mixer.music.get_busy() == 1:
        time.sleep(0.25)
        seconds += 0.25

    # PART 5
    mixer.quit()
    if delete_audio_file:
        os.remove(filename)
    print(f"audio file played for {seconds} seconds")


if __name__ == "__main__":
    say("You are worthy of love!")
    time.sleep(1.5)
    say("This is the main module. This program has finished. Thanks for running it!")
    """

# Import the required module for text 
# to speech conversion
from gtts import gTTS
  
# This module is imported so that we can 
# play the converted audio
import os
  
# The text that you want to convert to audio
mytext = 'Welcome to geeksforgeeks!'
  
# Language in which you want to convert
language = 'en'
  
# Passing the text and language to the engine, 
# here we have marked slow=False. Which tells 
# the module that the converted audio should 
# have a high speed
myobj = gTTS(text=mytext, lang=language, slow=False)
  
# Saving the converted audio in a mp3 file named
# welcome 
myobj.save("welcome.mp3")
  
# Playing the converted file
os.system("mpg321 welcome.mp3")