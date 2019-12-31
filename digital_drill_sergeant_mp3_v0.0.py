# -*- coding: utf-8 -*-
"""
Created on Mon Feb 25 08:18:27 2019

@author: Thinker
"""

import time # for spacing out exercises in time
from calendar import day_name
import pandas as pd # spreadsheet to dataframe
from random import randint # for randomizing left/right
from numpy import isnan # for checking if reps is 'not a number' (null)
from numpy import random # randomize times for stop go
from datetime import date # for noting date of run
#from pygame import mixer # for playing bell sound
from gtts import gTTS # google text to speech (allows mp3 output)
from pydub import AudioSegment # for importing gtts mp3 and concatenating audio

path = 'C:\\Users\\Thinker\\OneDrive - Bradley Analytics\\py\\ex\\'
mp3_path = path + 'mp3\\'
bell_path = path+'bell\\bell.mp3'

def play_bell(bell_path, audio_main):
    
    audio_new = AudioSegment.from_mp3(bell_path)
    audio_main = audio_main + audio_new
    
    return audio_main

def get_is_left():
    # randomly calculates boolean true or false
    # True means that 'left' side will go first
    # False means that 'right' side will go first
    
    is_left = bool(randint(0,1))
    return is_left

def get_sheetname(xlfilename):
    # pulls a list of sheet names from the excel file
    # displays sheet names along with index
    # allows user to enter index to specify sheet to use
    xl = pd.ExcelFile(xlfilename)
    sheet_names = xl.sheet_names
    
    # display index and sheet name
    sheet_index = 0
    for sheet in sheet_names:
        print(str(sheet_index) + ' - ' + sheet)
        sheet_index = sheet_index+1
    # reads in user input and convert to integer
    sheet_num = int(input('Please enter sheet number: \n'))
    # gets sheet_name that corresponds to sheet_num
    sheet_name = sheet_names[sheet_num]
    print('You entered ' + str(sheet_num) + ' - ' + sheet_name)
    return(sheet_name)

def get_today():
    # used to override day if I want to do a different program for today
    day_override = input('Please enter day number (0=Monday): \n')
    print('You entered: ' + day_override)
    try: # if try works
        day_override = int(day_override) # try converting to integer
        # check if between 0-6 inclusive
        valid = (day_override >= 0) & (day_override <=6)
        if valid == False:
            raise ValueError('Day Number is not between 0-6.')
        else:
            day_override = (day_override)
            today = day_name[day_override] # user input becomes today
    except:
        today = date.today().strftime('%A') # gets week day name 
    print('Running program for ' + today + '...')
    return(today)

def speak_and_wait(words, t, audio_main):
    
    # speak_and_wait(ex, 1, audio_main)
    
    # generate text to speech in english
    ggl_tts = gTTS(words, 'en')
    # save text to speech to mp3
    ggl_tts.save(mp3_path + 'temp_speech.mp3')
    
    # pydub imports mp3
    audio_new = AudioSegment.from_mp3(mp3_path + 'temp_speech.mp3')
    # create t seconds of silence
    audio_silence = AudioSegment.silent(t * 1000)

    # concatenate main audio with new audio that was just created with silence
    audio_main = audio_main + audio_new + audio_silence
    
    audio_main.export(mp3_path + 'drillsergeant.mp3', format='mp3')
    
    return audio_main
    
def stop_go(audio_main):
    
    go_times = random.random_sample(25)*15
    freeze_times = random.random_sample(25)*15
    for go, freeze in zip(go_times,freeze_times):
        audio_main = speak_and_wait("green light", go, audio_main)
        audio_main = speak_and_wait("red light", freeze, audio_main)
    audio_main = speak_and_wait("finished", 1, audio_main)

def log_time_elapsed_since(start_time, sheet_name):
        end_time = time.time()
        dur_mins = round((end_time - start_time)/60,2)
        with open(path + 'duration.csv', 'a') as f:
            f.write(sheet_name+','+str(dur_mins)+','+str(date.today())+'\n')

def crawl_sheet(is_left, sheet_name, today, xlfilename, bell_path, audio_main):
    
    df = pd.read_excel(xlfilename, sheet_name=sheet_name, engine = 'xlrd')
    df = df[df[today] == 'x'] # filters dataframe on records with 'x' for today
    for i in df.index:
        # if exercise is marked to be included, call it out, otherwise skip it
        if df.loc[i,'Include']=="x":
            # blank 'Reps' becomes nan once I assign to reps variable
            # assign values in row i to variables
            grp = df.loc[i,'Group']
            reps = df.loc[i, 'Reps']
            # replace null with '' string
            # reps is double by default.  int truncates "point 0" speech
            # str() makes it readable by tts
            reps_speech = '' if isnan(reps) else str(int(reps))
            
            ex = df.loc[i,'Exercise']
            opt = df.loc[i,'Option']
            lead_t = df.loc[i,'LeadTime']
            ex_t = df.loc[i,'ExTime']
            rest_t = df.loc[i,'RstTime']

            if grp == 'Intro':
                audio_main = speak_and_wait(ex, 1, audio_main)
                phrase = 'begin in ' + str(int(lead_t)) + ' seconds.'
                audio_main = speak_and_wait(phrase, lead_t, audio_main)
            elif grp == 'Outro':
                audio_main = speak_and_wait(ex, 1, audio_main)
            else:
                # lead_t is lead time to prepare for next exercise
                if lead_t > 0:
                    phrase = 'begin '+ex+' in '+str(int(lead_t))+' seconds'
                    audio_main = speak_and_wait(phrase, lead_t, audio_main)
                # L/R option will speak exercise twice, once for left and
                # once for right and will randomly select left, right order
                if opt == 'L/R':
                    if is_left == True:
                        first = 'left'
                        second = 'right'
                    else:
                        first = 'right'
                        second = 'left'
                    # if 'right' or 'left' should be in the middle of phrase
                    # substitute it for 'xxxxx'
                    if 'xxxxx' in ex:
                        speech1 = ex.replace('xxxxx', first)
                        speech2 = ex.replace('xxxxx', second)
                    else:
                        speech1 = ex + ' ' + first
                        speech2 = ex + ' ' + second
                    audio_main = speak_and_wait(speech1, ex_t, audio_main)
                    audio_main = speak_and_wait('other side', ex_t, audio_main)
                elif opt == 'stop go':
                    stop_go(audio_main)
                else:
                    audio_main = (speak_and_wait(reps_speech + ' ' + ex,
                                                 ex_t, 
                                                 audio_main))
                if grp == 'Breathing/Meditation':
                    audio_main = play_bell(bell_path, audio_main)
                    
                # rest_t is rest time after current exercise
                if rest_t > 0:
                    phrase = 'finished. rest for '+str(int(rest_t))+' seconds'
                    audio_main = speak_and_wait(phrase, rest_t, audio_main)
                    
    return audio_main

def main():
    
    # declare variables
    xlfilename = path+'ex_v2.xlsx'
    
    # create blank audiosegment - will contain whole program then output to mp3
    audio_main = AudioSegment.silent(0)
    
    # randomize left/right order of exercises for whole run
    is_left = get_is_left()
    
    # request user input to specify sheetname
    sheet_name = get_sheetname(xlfilename)
    
    # get day name of today for specifying daily routine
    today=get_today()
    
    # get start_time for loggin elapsed time
    start_time = time.time()
    
    # crawl through spreadsheet
    audio_main = (crawl_sheet(is_left, 
                              sheet_name, 
                              today, 
                              xlfilename, 
                              bell_path, 
                              audio_main
                              )
                  )
    
    
    # get end time and calculate elapsed time
    log_time_elapsed_since(start_time, sheet_name)
    
    audio_main.export(mp3_path + 'drillsergeant.mp3', format='mp3')

if __name__ == '__main__':
    main()

