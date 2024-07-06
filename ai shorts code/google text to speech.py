from pickle import TRUE
from gtts import gTTS
import openpyxl
import subprocess
import os
from pydub import AudioSegment

#run this with debugging. only then it will work properly

no_vids = 4
path = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\ai shorts code\\ai shorts psychology.xltx"

# To open the workbook 
# workbook object is created 
wb_obj = openpyxl.load_workbook(path) 

# Get workbook active sheet object 
# from the active attribute 
sheet_obj = wb_obj.active 

def create_text_to_speech():
    for i in range(2, 2+no_vids):
        for j in range(2, 4):
            cell_obj = sheet_obj.cell(row=i, column=j)
            quote = str(cell_obj.value)
            number_string1 = str(i - 1)
            number_string2 = str(j-1)
            tts = gTTS(quote, tld="com.au", slow = TRUE)
            output_path = os.path.join("C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\text to speech output", f"{number_string1}pt{number_string2}.mp3")
            tts.save(output_path)

def trim_tts_audio():
    for i in range(1,1+no_vids):
        for j in range(1,3):
            I=str(i)
            J=str(j)
            # Load the audio file
            audio = AudioSegment.from_mp3("C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\text to speech output\\"+I+"pt"+J+".mp3")

            # Trim silence from the end (adjust silence_thresh and duration parameters as needed)
            trimmed_audio = audio.strip_silence(silence_thresh=-50)
            trimmed_audio.export("C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\trimmed audio\\"+I+"pt"+J+".mp3", format='mp3')

def change_audio_speed():
    for i in range(1, 1+no_vids):
        for j in range(1, 3):
            I = str(i)
            J = str(j)
            input_file = os.path.join("C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\trimmed audio", f"{I}pt{J}.mp3")
            output_file = os.path.join("C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\audio with speed changed", f"{I}pt{J}.mp3")
            desired_duration = 3


            # Get the duration of the input audio file
            ffprobe_cmd = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', input_file]
            duration = float(subprocess.check_output(ffprobe_cmd).strip())


            # Calculate the speed factor
            speed_factor = duration/desired_duration
            spd=str(speed_factor)
            command = [
                'ffmpeg',
                '-i', input_file,
                '-filter:a', "atempo="+spd,
                '-vn',
                output_file
            ]
            try:
                subprocess.run(command, check=True)
                print(f"Successfully changed speed for {input_file} and saved to {output_file}")
            except subprocess.CalledProcessError as e:
                print(f"Error occurred: {e}")
                print(f"Command: {command}")

def merge_audio_files():
    # Command to merge audio files
    for i in range(1,1+no_vids):
        I=str(i)
        input1 = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\audio with speed changed\\"+I+"pt1.mp3"
        input2 = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\audio with speed changed\\"+I+"pt2.mp3"
        output = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\combined audio\\"+I+".mp3"
        command = [
            'ffmpeg',
            '-i', input1,
            '-i', input2,
            '-filter_complex', '[0:a][1:a]concat=n=2:v=0:a=1[out]',
            '-map', '[out]',
            output
        ]

        # Run ffmpeg command
        subprocess.run(command, check=True)


def overlay_audio():
    for i in range(1,1+no_vids):
        I = str(i)
        video_file = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\video input\\"+I+".mp4"
        audio_file = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\combined audio\\"+I+".mp3"
        output_file = "C:\\Users\\sudha\\Documents\\GitHub\\ai-shorts-code\\video output\\"+I+".mp4"
        cmd = [
            'ffmpeg',
            '-i', video_file,
            '-i', audio_file,
            '-c:v', 'copy', 
            '-filter_complex', '[0:a][1:a]amix=inputs=2:duration=first[outa]',
            '-map', '0:v',
            '-map', '[outa]',
            output_file
        ]

        # Run the ffmpeg command
        subprocess.run(cmd, check=True)


create_text_to_speech()
trim_tts_audio()
change_audio_speed()
merge_audio_files()
overlay_audio()