import pandas as pd
import pyautogui
import time
import pyaudio
import speech_recognition as sr
import wave
import logging
import wave
import assemblyai as aai
import json
import os
from multiprocessing import Process, Queue
import tkinter as tk
from tkinter import messagebox
from moviepy.editor import AudioFileClip, concatenate_audioclips
import mysql.connector
from mysql.connector import Error
import pygame
pyautogui.FAILSAFE = False
# Global constants and variables
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 44100
CHUNK_DURATION = 5
CHUNK_DURATION_R = 5
WAV_OUTPUT_DIR = "audio-chunks"
WAV_OUTPUT_DIR_R = "audio-chunks-records"
MAX_CHUNKS = 10
MAX_CHUNKS_R = 5
file_path = 'Sample File.xlsx'
sheet_name = 'Sheet1'
settings_file = 'settings.txt'
key = 'W8uA2ZszYIbe5NiVsluHq3aq7sIo0bSY'
final_text = ''

lists_dict = {
    "Voicemail": ["leave", "message", "been", "forward", "forwarded", "record", "please", "read"],
    "Disconnected": ["not", "available", "service", "disconnected"],
    "Connected": ["hello"],
}


class SingletonTk(tk.Tk):
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(SingletonTk, cls).__new__(cls, *args, **kwargs)
        return cls._instance
    
def find_matching_list(original_string, lists_dict):
    original_string = original_string.lower()
    for list_name, word_list in lists_dict.items():
        word_list = [word.lower() for word in word_list]
        if any(word in original_string for word in word_list):
            return list_name  
    return "Ringing"

def save_chunk_as_wav(frames, filename):
    with wave.open(filename, 'wb') as wf:
        wf.setnchannels(CHANNELS)
        wf.setsampwidth(pyaudio.PyAudio().get_sample_size(FORMAT))
        wf.setframerate(RATE)
        wf.writeframes(b''.join(frames))

def transcribe_audio(file_path):
    recognizer = sr.Recognizer()
    try:
        with sr.AudioFile(file_path) as source:
            audio_data = recognizer.record(source)
            text = recognizer.recognize_google(audio_data)
            return text
    except sr.UnknownValueError:
        return " "
    except sr.RequestError as e:
        return f"service; {e}"


def stream_and_transcribe_record(queue):
    audio = pyaudio.PyAudio()
    stream = audio.open(format=FORMAT,
                        channels=CHANNELS,
                        rate=RATE,
                        input=True,
                        frames_per_buffer=CHUNK)
    
    print("Record...")
    frames = []
    chunk_frames = []
    final_text = ''
    chunk_count = 1
    play = False
    try:
        while chunk_count <= MAX_CHUNKS:
            data = stream.read(CHUNK, exception_on_overflow=False)
            frames.append(data)
            chunk_frames.append(data)
            chunk_size = RATE * CHUNK_DURATION_R * audio.get_sample_size(FORMAT)
            
            if len(b''.join(chunk_frames)) >= chunk_size:
                chunk_filename = os.path.join(WAV_OUTPUT_DIR_R, f"chunk_{chunk_count}.wav")
                save_chunk_as_wav(chunk_frames, chunk_filename)
                print("record")
                chunk_frames = [] 
                chunk_count += 1
                
    except KeyboardInterrupt: 
        print("Stopped by user.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        stream.stop_stream()
        stream.close()
        audio.terminate()
        print("Audio stream closed.")
        queue.put(None)  # Indicate that transcription is done
        return final_text
    

def stream_and_transcribe(queue):
    audio = pyaudio.PyAudio()
    stream = audio.open(format=FORMAT,
                        channels=CHANNELS,
                        rate=RATE,
                        input=True,
                        frames_per_buffer=CHUNK)
    
    print("Translating...")
    frames = []
    chunk_frames = []
    final_text = ''
    chunk_count = 1
    play = False
    break_ = False
    try:
        while chunk_count <= MAX_CHUNKS:
            if break_:
                print("Breaking loop.")
                continue
            data = stream.read(CHUNK, exception_on_overflow=False)
            frames.append(data)
            chunk_frames.append(data)
            chunk_size = RATE * CHUNK_DURATION_R * audio.get_sample_size(FORMAT)
            
            if len(b''.join(chunk_frames)) >= chunk_size:
                chunk_filename = os.path.join(WAV_OUTPUT_DIR_R, f"chunk_{chunk_count}.wav")
                # save_chunk_as_wav(chunk_frames, chunk_filename)
                text = transcribe_audio(chunk_filename)
                # aai.settings.api_key = "767c81af5ca34f069ed6700d0aa9ec19"
                # transcriber = aai.Transcriber()
                # transcript = transcriber.transcribe(chunk_filename)
                # text = transcript.text
                result = find_matching_list(text, lists_dict)
                print(result)
                if result == "Connected" and not play:
                    print("Playing alert sound...")
                    play = True
                    pygame.mixer.init()
                    pygame.mixer.music.load('audio.mp3')
                    pygame.mixer.music.play()
                    while pygame.mixer.music.get_busy(): 
                        pygame.time.Clock().tick(10)
                if result == "Voicemail":
                    break_ = True
                    chunk_count = MAX_CHUNKS
                if result == "Connected":
                    break_ = True
                    chunk_count = MAX_CHUNKS

                chunk_frames = [] 
                chunk_count += 1
                
    except KeyboardInterrupt: 
        print("Stopped by user.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        stream.stop_stream()
        stream.close()
        audio.terminate()
        print("Audio stream closed.")
        queue.put(None)  # Indicate that transcription is done
        return final_text

def read_wav():
    final_text = ''
    chunk_count = 1
    while chunk_count <= MAX_CHUNKS:
        chunk_filename = os.path.join(WAV_OUTPUT_DIR_R, f"chunk_{chunk_count}.wav")
        if os.path.exists(chunk_filename):
            aai.settings.api_key = "767c81af5ca34f069ed6700d0aa9ec19"
            transcriber = aai.Transcriber()
            transcript = transcriber.transcribe(chunk_filename)
            text= transcript.text
            # text = transcribe_audio(chunk_filename)
            final_text += text + " "
            print(f"{chunk_filename}: {text}")
        else:
            print(f"Warning: {chunk_filename} does not exist.")
        chunk_count += 1
        time.sleep(1)
    return final_text.strip()

def merge_chunks_to_mp3(output_filename, num_chunks=4, input_dir=WAV_OUTPUT_DIR_R):
    audio_clips = []
    for i in range(1, num_chunks + 1):
        chunk_filename = os.path.join(input_dir, f"chunk_{i}.wav")
        if os.path.exists(chunk_filename):
            audio_clip = AudioFileClip(chunk_filename)
            audio_clips.append(audio_clip)
        else:
            print(f"Warning: {chunk_filename} does not exist.")
    
    if audio_clips:
        final_audio = concatenate_audioclips(audio_clips)
        final_audio.write_audiofile(os.path.join(input_dir, output_filename), codec='mp3')
        print(f"Successfully merged {num_chunks} chunks into {output_filename}")
    else:
        print("No valid chunks were found to merge.")

def load_settings():
    try:
        with open(settings_file, 'r') as file:
            lines = file.readlines()
            settings = {
                'focus_x': int(lines[0].strip()),
                'focus_y': int(lines[1].strip()),
                'initiate_x': int(lines[2].strip()),
                'initiate_y': int(lines[3].strip()),
                'decline_x': int(lines[4].strip()),
                'decline_y': int(lines[5].strip()),
                'last_index': int(lines[6].strip()),
                'merge_chunks': lines[7].strip() == 'Yes'
            }
    except FileNotFoundError:
        settings = {
            'focus_x': 0,
            'focus_y': 0,
            'initiate_x': 0,
            'initiate_y': 0,
            'decline_x': 0,
            'decline_y': 0,
            'last_index': 0,
            'merge_chunks': False
        }
    return settings

def save_settings(focus_x, focus_y, initiate_x, initiate_y, decline_x, decline_y, last_index, merge_chunks):
    with open(settings_file, 'w') as file:
        file.write(f"{focus_x}\n")
        file.write(f"{focus_y}\n")
        file.write(f"{initiate_x}\n")
        file.write(f"{initiate_y}\n")
        file.write(f"{decline_x}\n")
        file.write(f"{decline_y}\n")
        file.write(f"{last_index}\n")
        file.write("Yes" if merge_chunks else "No")

def analyze_transcription(text):
    keywords = ["yes", "no"]
    text = text.lower().strip()
    
    for keyword in keywords:
        if keyword in text:
            return keyword
    return "NAN"

def start_automation():
    file_path = 'Sample File.xlsx'
    sheet_name = 'Sheet1'
    
    focus_x = int(focus_number_input_x.get())
    focus_y = int(focus_number_input_y.get())
    initiate_x = int(initiate_call_x.get())
    initiate_y = int(initiate_call_y.get())
    decline_x = int(decline_call_x.get())
    decline_y = int(decline_call_y.get())
    merge_chunks = merge_chunks_var.get() == "Yes"
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df['Result'] = df['Result'].astype(str)
    time.sleep(2)
    pyautogui.hotkey("alt", "tab")
    should_break = False
    count = 0

    settings = load_settings()
    last_index = settings['last_index']
    for index, row in df.iterrows():
        if should_break:
            break 

        if index <= last_index:
            continue

        try:
            connection = mysql.connector.connect(
                host='165.73.249.12',
                user='waqeel786',
                password='waqeel3726487',
                database='licence'
            )

            if connection.is_connected():
                cursor = connection.cursor()
                select_query = "SELECT `limit` FROM `user` WHERE `key` = %s"
                cursor.execute(select_query, (key,))
                result = cursor.fetchone()
                
                if result:
                    current_limit = result[0]
                    
                    if current_limit > 0:
                        update_query = """
                        UPDATE `user`
                        SET `limit` = CASE
                            WHEN `limit` > 0 THEN `limit` - 1
                            ELSE `limit`
                        END
                        WHERE `key` = %s
                        """
                        cursor.execute(update_query, (key,))
                        connection.commit()
                    else:
                        messagebox.showwarning("API Key Limit", "API key expired. Please renew your key.")
                        should_break = True
                else:
                    messagebox.showwarning("API Key Limit", "API key expired. Please renew your key.")
                    should_break = True

        except Error as e:
            messagebox.showwarning("API Key Limit", "API key expired. Please renew your key.")
            should_break = True

        if connection and connection.is_connected():
            connection.close()

        if should_break:
            break  

        time.sleep(1)
        pyautogui.click(focus_x, focus_y)
        time.sleep(1)
        pyautogui.write(f"{row['Number']}", interval=0.1)  
        time.sleep(1)
        pyautogui.click(initiate_x, initiate_y)

        queue_record = Queue()
        queue_transcribe = Queue()
        
        process1 = Process(target=stream_and_transcribe, args=(queue_transcribe,))
        process2 = Process(target=stream_and_transcribe_record, args=(queue_record,))
        process1.start()
        process2.start()
        
        process1.join()
        process2.join()
        
        # Ensure both processes have completed
        if queue_transcribe.get() is None and queue_record.get() is None:
            time.sleep(1)  
            pyautogui.click(decline_x, decline_y)
            text_call = read_wav()
            result = analyze_transcription(text_call)
            
            output_filename = f"{row['Number']}.mp3"
            if merge_chunks:
                merge_chunks_to_mp3(output_filename, num_chunks=MAX_CHUNKS_R, input_dir=WAV_OUTPUT_DIR_R)
            
            df.at[index, 'Result'] = result if result is not None else 'No Match'
            print(f"Processed {row['Number']}: {result}")

            count += 1
            last_index = index 
            df.to_excel(file_path, sheet_name=sheet_name, index=False)
            save_settings(focus_x, focus_y, initiate_x, initiate_y, decline_x, decline_y, last_index, merge_chunks)

    total_processed_label.config(text=f"Total Processed: {count}")
    last_index_label.config(text=f"Last Processed Index: {last_index}")
    
logging.basicConfig(filename='app.log', level=logging.DEBUG)

def setup_gui():
    logging.debug("Setting up GUI...")
    global focus_number_input_x, focus_number_input_y, initiate_call_x, initiate_call_y
    global decline_call_x, decline_call_y, start_button, total_processed_label
    global last_index_label, merge_chunks_var, root

    if 'root' in globals() and root is not None and root.winfo_exists():
        return  

    root = SingletonTk()
    root.title("Ringcentral Automation Tool - Powered By : Trestech")
    settings = load_settings()

    focus_x = settings['focus_x']
    focus_y = settings['focus_y']
    initiate_x = settings['initiate_x']
    initiate_y = settings['initiate_y']
    decline_x = settings['decline_x']
    decline_y = settings['decline_y']
    merge_chunks = "Yes" if settings['merge_chunks'] else "No"

    tk.Label(root, text="Focus Number Input X:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    focus_number_input_x = tk.Entry(root)
    focus_number_input_x.grid(row=0, column=1, padx=5, pady=5)
    focus_number_input_x.insert(0, focus_x)
    tk.Label(root, text="Y:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
    focus_number_input_y = tk.Entry(root)
    focus_number_input_y.grid(row=0, column=3, padx=5, pady=5)
    focus_number_input_y.insert(0, focus_y)

    tk.Label(root, text="Initiate Call X:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    initiate_call_x = tk.Entry(root)
    initiate_call_x.grid(row=1, column=1, padx=5, pady=5)
    initiate_call_x.insert(0, initiate_x)
    tk.Label(root, text="Y:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
    initiate_call_y = tk.Entry(root)
    initiate_call_y.grid(row=1, column=3, padx=5, pady=5)
    initiate_call_y.insert(0, initiate_y)

    tk.Label(root, text="Decline Call X:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    decline_call_x = tk.Entry(root)
    decline_call_x.grid(row=2, column=1, padx=5, pady=5)
    decline_call_x.insert(0, decline_x)
    tk.Label(root, text="Y:").grid(row=2, column=2, padx=5, pady=5, sticky="w")
    decline_call_y = tk.Entry(root)
    decline_call_y.grid(row=2, column=3, padx=5, pady=5)
    decline_call_y.insert(0, decline_y)

    merge_chunks_var = tk.StringVar(value=merge_chunks)
    tk.Label(root, text="You Want Call Records?").grid(row=5, column=0, padx=10, pady=5, sticky="w")
    tk.Radiobutton(root, text="Yes", variable=merge_chunks_var, value="Yes").grid(row=5, column=1, padx=5, pady=5)
    tk.Radiobutton(root, text="No", variable=merge_chunks_var, value="No").grid(row=5, column=2, padx=5, pady=5)

    start_button = tk.Button(root, text="Start Automation", command=start_automation)
    start_button.grid(row=6, column=0, columnspan=4, padx=10, pady=10)

    total_processed_label = tk.Label(root, text="Total Processed: 0")
    total_processed_label.grid(row=3, column=0, columnspan=4, padx=10, pady=5)
    last_index_label = tk.Label(root, text="Last Processed Index: 0")
    last_index_label.grid(row=4, column=0, columnspan=4, padx=10, pady=5)

    root.mainloop()

setup_done = False
def main():
    global setup_done
    if not setup_done:
        setup_gui()
        setup_done = True
    else:
        print("GUI setup has already been done.")

main()