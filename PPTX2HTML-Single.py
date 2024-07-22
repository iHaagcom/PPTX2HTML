import os
import shutil
import win32com.client as win32
import zipfile
import webbrowser
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QProgressBar, QVBoxLayout, QLabel, QTextEdit, QHBoxLayout, QLineEdit, QMessageBox, QCheckBox, QSlider, QGroupBox, QComboBox
from PyQt5.QtCore import QThread, pyqtSignal, QObject, pyqtSlot, Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
import sys
import traceback
import logging
import time
from win32com.client import gencache
import pythoncom
import subprocess
import json
#Added checkbox for reboot

#logging.basicConfig(level=logging.DEBUG)
#logger = logging.getLogger(__name__)

logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)
def custom_excepthook(exc_type, exc_value, exc_traceback):
    # Get the last line of the traceback
    tb_last = traceback.extract_tb(exc_traceback)[-1]
    filename = tb_last.filename.split('\\')[-1]  # Get just the filename, not the full path
    line_num = tb_last.lineno
    logging.error(f"{exc_type.__name__}: {exc_value} (in {filename}, line {line_num})")

sys.excepthook = custom_excepthook

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

USE_PSEXEC = True
PSEXEC_PATH = resource_path("psshutdown.exe") 
ffmpeg_path = "ffmpeg.exe"
whisper_exe = "Whisper\\main.exe"
model_path = "Whisper\\ggml-model-whisper-base.bin"
#whisper\main.exe -m .\Whisper\ggml-model-whisper-base.bin -f output.wav -ml 500 -otxt -ovtt -osrt
#ffmpeg.exe -i media1.m4a -acodec pcm_s16le -ar 16000 -ac 1 test.wav
#Whisper\stream.exe -m whisper\ggml-model-whisper-base.bin

def check_ffmpeg_path():
    return os.path.exists(ffmpeg_path)

def check_whisper_exe():
    return os.path.exists(whisper_exe)

def check_7zip():
    return os.path.exists("C:\\Program Files\\7-Zip\\7z.exe")
## Added a reset command ##
CUSTOM_MAPPINGS = {
    r'\\CustomPlaceHolder': {
        "machine": "192.168.0.100",
        "username": "username_changeme",
        "password": "password_changeme"
    }
}
def send_reset_command(export_path, machine, reboot=False):
    config = CUSTOM_MAPPINGS.get(export_path, {})
    target_machine = config.get('machine', machine)
    username = config.get('username')
    password = config.get('password')

    if USE_PSEXEC:
        return send_reset_command_psexec(target_machine, username, password, reboot)
    else:
        return send_reset_command_shutdown(target_machine, username, password, reboot)

def send_reset_command_shutdown(machine, username=None, password=None, reboot=False):
    if reboot:
        cmd = ['shutdown', '/r', '/t', '0', '/m', f'\\\\{machine}']
    else:
        cmd = ['shutdown', '/s', '/t', '0', '/m', f'\\\\{machine}']
    
    if username and password:
        cmd.extend(['/u', username, '/p', password])
    
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        action = "Reboot" if reboot else "Shutdown"
        logging.info(f"{action} command sent to {machine}")
        return True
    except subprocess.CalledProcessError as e:
        action = "reboot" if reboot else "shutdown"
        logging.error(f"Failed to send {action} command to {machine}: {e.stderr}")
        return False

def send_reset_command_psexec(machine, username=None, password=None, reboot=False):
    cmd = [PSEXEC_PATH, f'\\\\{machine}']
    if username and password:
        cmd.extend(['-u', f'{machine}\\{username}', '-p', password])
    cmd.extend(['-f', '-t'])
    if reboot:
        cmd.append('-r')
    cmd.extend(['shutdown', '/f', '/t', '0'])
    
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        action = "Reboot" if reboot else "Shutdown"
        logging.info(f"{action} command sent to {machine} using PSExec")
        return True
    except subprocess.CalledProcessError as e:
        action = "reboot" if reboot else "shutdown"
        logging.error(f"Failed to send {action} command to {machine} using PSExec: {e.stderr}")
        return False

#Check if FFmpeg is accessible
#try:
#    result = subprocess.run([ffmpeg_path, "-version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
#    logging.info("FFmpeg is accessible and executable.")
#except subprocess.CalledProcessError as e:
#    logging.error(f"FFmpeg is not accessible or executable: {e}")
#except PermissionError as pe:
#    logging.error(f"Permission error: {pe}. Please ensure you have the necessary permissions to run the FFmpeg executable.")


def ensure_powerpoint_closed():
    try:
        subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info("Forcefully closed any existing PowerPoint instances")
    except subprocess.CalledProcessError:
        logging.debug("No PowerPoint instances were running")

def extract_media_from_pptx(pptx_file, output_dir):
    try:
        with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                try:
                    zip_ref.extract(file_info, output_dir)
                except zipfile.BadZipFile as e:
                    logging.info(f"CRC error detected for file {file_info.filename}. Attempting extraction with 7-Zip.")
                    extract_with_7zip(pptx_file, output_dir, file_info.filename)
                except Exception as e:
                    logging.info(f"Error extracting file {file_info.filename}: {e}")
    except Exception as e:
        logging.info(f"Error extracting media from PPTX: {e}")
    
    media_dir = os.path.join(output_dir, 'ppt', 'media')
    if not os.path.exists(media_dir):
        logging.info(f"No media directory found in {pptx_file}")
        return []

    media_files = [os.path.join(media_dir, f) for f in os.listdir(media_dir)]
    return media_files

def extract_with_7zip(pptx_file, output_dir, file_to_extract):
    try:
        output_path = os.path.join(output_dir, file_to_extract)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        result = subprocess.run(['C:\\Program Files\\7-Zip\\7z.exe', 'e', pptx_file, f'-o{output_dir}', file_to_extract], capture_output=True, text=True)
        if result.returncode == 0:
            logging.info(f"Extracted {file_to_extract} with 7-Zip to {output_path}")
            return True
        else:
            logging.info(f"Error extracting with 7-Zip: {result.stderr}")
    except Exception as e:
        logging.info(f"Error extracting with 7-Zip: {e}")
    return False

def check_and_replace_zero_byte_media(output_dir):
    media_path = os.path.join(output_dir, 'ppt', 'media', 'media1.m4a')
    if os.path.exists(media_path) and os.path.getsize(media_path) == 0:
        logging.info(f"File {media_path} is zero bytes. Checking for alternative extraction.")
        alternative_path = os.path.join(output_dir, 'media1.m4a')
        if os.path.exists(alternative_path) and os.path.getsize(alternative_path) > 0:
            logging.info(f"Replacing zero-byte file with {alternative_path}")
            shutil.move(alternative_path, media_path)
        else:
            logging.info("No valid alternative file found.")

def segment_transcription(whisper_output, slide_durations):
    if not whisper_output:
        return []
    
    lines = whisper_output.strip().split('\n')
    segments = []
    
    for line in lines:
        if '[' in line and ']' in line:
            time_range, text = line.split(']', 1)
            start_time, end_time = map(parse_whisper_time, time_range.strip('[').split(' --> '))
            segments.append({
                'start': start_time,
                'end': end_time,
                'text': text.strip()
            })
    
    return segments

def parse_whisper_time(time_str):
    h, m, s = time_str.split(':')
    return int(h) * 3600 + int(m) * 60 + float(s)

def convert_to_wav(audio_path):
    if not os.path.exists(audio_path):
        logging.info(f"Audio file not found: {audio_path}")
        return None
    
    # Check if the file is already a WAV file
    if audio_path.lower().endswith('.wav'):
        logging.info(f"File is already a WAV file: {audio_path}")
        return audio_path
    
    wav_path = os.path.splitext(audio_path)[0] + '.wav'
    if os.path.exists(wav_path):
        logging.info(f"WAV file already exists: {wav_path}")
        return wav_path
    
    logging.info(f"Converting {audio_path} to {wav_path}")
    
    command = [
        ffmpeg_path,
        "-i", audio_path,
        "-acodec", "pcm_s16le",
        "-ar", "16000",
        "-ac", "1",
        "-f", "wav",
        wav_path
    ]
    
    try:
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        logging.info(f"Conversion successful: {wav_path}")
    except subprocess.CalledProcessError as e:
        logging.info(f"FFmpeg conversion failed: {e}")
        logging.info(f"FFmpeg stderr: {e.stderr.decode('utf-8')}")
        return None
    
    if os.path.exists(wav_path):
        return wav_path
    else:
        logging.info(f"WAV file not found after conversion: {wav_path}")
        return None

def transcribe_audio(self, audio_path):
        if not audio_path or not os.path.exists(audio_path):
            logging.info(f"Audio file not found for transcription: {audio_path}")
            return None

        txt_path = audio_path + ".txt"
        if os.path.exists(txt_path):
            with open(txt_path, 'r', encoding='utf-8') as f:
                return f.read().strip()

        command = [
            whisper_exe,
            "-m", model_path,
            "-f", audio_path,
            "--output-txt"
        ]
        
        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            logging.info(f"Whisper output: {result.stdout}")
            return result.stdout.strip()
        except subprocess.CalledProcessError as e:
            logging.info(f"Error transcribing audio: {e}")
            logging.info(f"Stderr: {e.stderr}")
            return None

def export_as_jpg_impl(template_file, export_location, slide_duration, progress_callback=None, transcribe_audio_func=None):
    slide_list = []
    ppt_app = None
    presentation = None
    global_audio_path = None
    global_transcription = ""
    whisper_output = global_transcription  # Assume global_transcription contains Whisper output
    try:
        pythoncom.CoInitialize()
        ppt_app = win32.dynamic.Dispatch("PowerPoint.Application")
        presentation = ppt_app.Presentations.Open(template_file, ReadOnly=True, Untitled=False, WithWindow=False)
        logging.info("Exporting presentation to JPG")
        presentation.Export(export_location, "JPG")
        time.sleep(1)
        logging.info("Beginning to process slides")
        slides = presentation.Slides

        media_files = extract_media_from_pptx(template_file, export_location)
        media_files_dict = {os.path.basename(f): f for f in media_files}

        # Process audio once for all slides
        for media_file in media_files:
            if media_file.lower().endswith(('.m4a', '.wav', '.mp3')):
                global_audio_path = media_file
                check_and_replace_zero_byte_media(export_location)  # Check and replace zero-byte media file
                if transcribe_audio_func:
                    wav_path = convert_to_wav(global_audio_path)
                    if wav_path:
                        global_transcription = transcribe_audio_func(wav_path) or ""
                        logging.info(f"Generated transcription: {global_transcription}")
                    else:
                        logging.error(f"Skipping transcription due to failed conversion: {global_audio_path}")
                break  # Process only one audio file

        slide_durations = []
        for count, slide in enumerate(slides, start=1):
            slide_timer = slide.SlideShowTransition.AdvanceTime
            slide_durations.append(slide_duration if slide_timer == 0 else slide_timer)

        # Segment the transcription based on slide durations
        transcription_segments = segment_transcription(global_transcription, slide_durations) if transcribe_audio_func else []

        for count, slide in enumerate(slides, start=1):
            if progress_callback:
                progress = 20 + (50 * count / len(slides))
                progress_callback(int(progress))
            slide_timer = slide.SlideShowTransition.AdvanceTime
            slide_path = os.path.join(export_location, f"Slide{count}.JPG")
            video_path = None
            video_duration = 0
            media_exported = False

            for shape_index in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes.Item(shape_index)
                try:
                    if shape.Type == 16:
                        for media_file_name, media_file_path in media_files_dict.items():
                            if media_file_name.startswith(f"media") and media_file_path.lower().endswith('.mp4'):
                                video_path = os.path.relpath(media_file_path, export_location)
                                logger.info(f"Media exported for slide {count}: {video_path}")
                                slide_path = os.path.relpath(media_file_path, export_location)
                                video_duration = shape.MediaFormat.Length / 1000 if hasattr(shape.MediaFormat, 'Length') else 5
                                media_exported = True
                                del media_files_dict[media_file_name]
                                break
                            elif media_file_name.startswith(f"media") and media_file_path.lower().endswith(('.m4a', '.wav', '.mp3')):
                                audio_path = os.path.relpath(media_file_path, export_location)
                                logger.info(f"Audio exported for slide {count}: {audio_path}")
                                del media_files_dict[media_file_name]
                                break

                except Exception as e:
                    logger.error(f"Error processing shape on slide {count}: {e}")

            if not media_exported:
                slide_list.append({
                    'file_path': slide_path,
                    'timer': slide_duration if slide_timer == 0 else slide_timer,
                    'audio_path': os.path.relpath(global_audio_path, export_location) if global_audio_path else None,
                    'video_path': video_path,
                    'slide_has_audio': global_audio_path is not None or video_path is not None,
                    #'transcription': transcription_segments[count - 1] if count - 1 < len(transcription_segments) else ""
                    'transcription': transcription_segments[count - 1] if transcribe_audio_func and count - 1 < len(transcription_segments) else ""
                })
            else:
                slide_list.append({
                    'file_path': slide_path if not video_path else video_path,
                    'timer': video_duration,
                    'audio_path': os.path.relpath(global_audio_path, export_location) if global_audio_path else None,
                    'video_path': video_path,
                    'slide_has_audio': global_audio_path is not None or video_path is not None,
                    #'transcription': transcription_segments[count - 1] if count - 1 < len(transcription_segments) else ""
                    'transcription': transcription_segments[count - 1] if transcribe_audio_func and count - 1 < len(transcription_segments) else ""
                })

        presentation.Close()
        ppt_app.Quit()
    except Exception as ex:
        logging.info(f"Error: {ex}")
        #import traceback
        #traceback.logging.info_exc()

    return slide_list

_HTMLTemplate = """
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <title>Powerpoint Presentation</title>
    <meta http-equiv="content-type" content="text/html; charset=utf-8" />
    <meta http-equiv="cache-control" content="max-age=0" />
    <meta http-equiv="cache-control" content="no-cache" />
    <meta http-equiv="expires" content="-1" />
    <meta http-equiv="pragma" content="no-cache" />
</head>
<script>
var audioElement = null;
var videoElement = null;
var slideDuration = %DURATION%;
var isImage = %ISIMAGE%;
var slideHasAudio = %SLIDEHASAUDIO%;
var globalAudioPath = "%GLOBALAUDIOPATH%";
var timer = null;
var transcription = %TRANSCRIPTION%;
var slideIndex = %SLIDEINDEX%;
var totalSlides = %TOTALSLIDES%;
var useWhisper = %USEWHISPER%;

function NextPage(){
    if (videoElement) {
        videoElement.pause();
    }
    if (audioElement) {
        if (slideIndex == totalSlides) {
            // If this is the last slide, reset audio time to 0
            sessionStorage.setItem('audioCurrentTime', 0);
        } else {
            sessionStorage.setItem('audioCurrentTime', audioElement.currentTime);
        }
    }
    clearTimeout(timer);
    window.location.href = "./%NEXTPAGE%";
}

function startTimer() {
    console.log("Starting timer for " + slideDuration + " seconds");
    clearTimeout(timer);
    timer = setTimeout(NextPage, slideDuration * 1000);
}

function handleAudio() {
    audioElement = document.getElementById("bg-audio");
    if (audioElement) {
        var storedTime = parseFloat(sessionStorage.getItem('audioCurrentTime')) || 0;
        audioElement.currentTime = storedTime;
        audioElement.play().catch(error => {
            console.log("Audio autoplay prevented:", error);
            startTimer();
            });
        audioElement.onended = function() {
            if (slideIndex == totalSlides) {
                sessionStorage.setItem('audioCurrentTime', 0);
            }
        };
        audioElement.ontimeupdate = updateTranscription;
    } else {
        console.log("No audio element found");
        startTimer();
    }
}

function handleMedia() {
    if (isImage) {
        console.log("This is an image slide");
        startTimer();
    } else {
        videoElement = document.querySelector("video");
        if (videoElement) {
            console.log("This is a video slide");
            videoElement.onended = NextPage;
            videoElement.muted = true;  // Mute video to allow autoplay
            videoElement.play().then(() => {
                console.log("Video started playing");
                try {
                    videoElement.muted = false;  // Try to unmute
                } catch (error) {
                    console.log("Unmuting failed:", error);
                    videoElement.muted = true;
                }
            }).catch(function(error) {
                console.log("Video autoplay was prevented:", error);
                // Move to next slide after a delay if video can't play
                videoElement.muted = false;
                videoElement.play();
                setTimeout(NextPage, slideDuration * 1000);  // 5 second delay, adjust as needed
            });
        // Add event listeners for debugging
            videoElement.addEventListener('loadedmetadata', () => {
                console.log("Video metadata loaded");
            });
            videoElement.addEventListener('canplay', () => {
                console.log("Video can play");
            });
            videoElement.addEventListener('playing', () => {
                console.log("Video is playing");
            });
            videoElement.addEventListener('pause', () => {
                console.log("Video paused");
                videoElement.muted = true;  // Mute video to allow autoplay
            videoElement.play();
            });
            videoElement.addEventListener('error', (e) => {
                console.error("Video error:", e);
            });
        } else {
            console.log("No video element found");
            startTimer();
        }
    }
}

function updateTranscription() {
    if (!useWhisper || !transcription) {
    document.getElementById("transcription").style.display = "none";
        return
    }
    var currentTime = audioElement ? audioElement.currentTime : 0;
    var transcriptionElement = document.getElementById("transcription");
    
    if (transcription.start <= currentTime && currentTime <= transcription.end) {
        transcriptionElement.textContent = transcription.text;
        transcriptionElement.style.display = "block";
    } else {
        transcriptionElement.textContent = "";
        transcriptionElement.style.display = "none";
    }
}

window.onload = function() {
    console.log("Slide Index:", slideIndex);
    console.log("Total Slides:", totalSlides);
    console.log("Transcription data:", transcription);
    
    if (slideIndex == 1) {
        // If this is the first slide (index), check if we're coming from the last slide
        var lastPageVisited = sessionStorage.getItem('lastPageVisited');
        if (lastPageVisited == totalSlides) {
            // We've looped back to the start, so reset the audio time
            sessionStorage.setItem('audioCurrentTime', 0);
        }
    }
    
    // Store the current slide index
    sessionStorage.setItem('lastPageVisited', slideIndex);
    
    handleAudio();
    handleMedia();
    updateTranscription();
    if (useWhisper && transcription) {
        console.log("Initializing transcription");
        updateTranscription();
    }
};
</script>
<style type="text/css">
html, body {
    width: 100%;
    height: 100%;
    margin: 0;
    padding: 0;
    background-color: #000000;
}
img#bg {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    object-fit: contain;
}
video {
    width: 100%;
    height: 100%;
    object-fit: contain;
}
#transcription {
    position: absolute;
    bottom: 10px;
    left: 10px;
    right: 10px;
    color: white;
    background-color: rgba(0,0,0,0.5);
    padding: 5px;
    font-size: 18px;
    text-align: center;
}
</style>
<body>
    %CONTENT%
    %TRANSCRIPTION_DIV%
</body>
</html>
"""

def create_html(html_directory, slide_list, use_whisper):
    html_directory = os.path.abspath(html_directory)
    logging.info(f"Using absolute html_directory path: {html_directory}")
    if not slide_list:
        logging.warning("No slides to process.")
        return
    logging.info(f"Creating HTML for {len(slide_list)} slides")
    global_audio_path = next((slide['audio_path'] for slide in slide_list if slide['audio_path']), "")
    total_slides = len(slide_list)
    
    for i, slide in enumerate(slide_list, start=1):
        logging.info(f"Processing slide {i} of {total_slides}")
        content = ""
        file_name = "index.html" if i == 1 else f"Slide{i:02d}.html"
        next_file = "index.html" if i == len(slide_list) else f"Slide{i+1:02d}.html"
        logging.info(f"Creating file: {file_name}")

        is_image = not slide['video_path']

        if is_image:
            image_path = os.path.basename(slide['file_path'])
            content = f"<img src=\"{image_path}\" id=\"bg\">"
        else:
            video_path = slide['video_path']
            content = f"<video webkit-playsinline=\"true\" playsinline=\"true\" autoplay=\"\"><source src=\"{video_path}\" type=\"video/mp4\">Your browser does not support the video tag.</video>"

        if global_audio_path:
            content += f"<audio id=\"bg-audio\" src=\"{global_audio_path}\"></audio>"

        logging.info(f"Transcription for slide {i}:", slide['transcription'])  # Debug logging.info
        transcription_div = '<div id="transcription"></div>' if use_whisper else ''
        
        # Format the transcription data
        if use_whisper and slide['transcription']:
            transcription_data = {
                "start": slide['transcription'].get('start', 0),
                "end": slide['transcription'].get('end', slide['timer']),
                "text": slide['transcription'].get('text', '')
            }
            logging.info(f"Transcription data for slide {i}: {transcription_data}")
        else:
            transcription_data = None #{"start": 0, "end": 0, "text": ""}
        
        escaped_transcription = json.dumps(transcription_data) if transcription_data else 'null'
        
        text = _HTMLTemplate.replace("%CONTENT%", content)
        text = text.replace("%NEXTPAGE%", next_file)
        text = text.replace("%DURATION%", str(slide['timer']))
        text = text.replace("%ISIMAGE%", str(is_image).lower())
        text = text.replace("%SLIDEHASAUDIO%", str(slide['slide_has_audio']).lower())
        text = text.replace("%GLOBALAUDIOPATH%", global_audio_path)
        text = text.replace("%TRANSCRIPTION%", escaped_transcription)
        text = text.replace("%TRANSCRIPTION_DIV%", transcription_div)
        text = text.replace("%SLIDEINDEX%", str(i))
        text = text.replace("%TOTALSLIDES%", str(total_slides))
        text = text.replace("%USEWHISPER%", str(use_whisper).lower())
        logging.info("Finished creating all HTML files")
        with open(os.path.join(html_directory, file_name), "w", encoding='utf-8') as file:
            file.write(text)

EXPORT_LOCATIONS = {
    "Custom Placeholder": {
        r'\\custom\path': ("Custom Placeholder", "Custom Placeholder", "Custom Placeholder"),
    },
    "Custom Placeholder2": {
        r'\\custom\path': ("Custom Placeholder", "Custom Placeholder", "Custom Placeholder"),
    },
    "Local Paths": {
        #r'C:\Users\Public\Documents': ("Local Path", "Public Documents", "Local Path"),
        r'D:\Exports': ("Local Path", "Local Directory", "Local Path"),
    }
}

def open_html_in_browser(html_directory):
    index_file = os.path.join(html_directory, "index.html")
    if os.path.exists(index_file):
        webbrowser.open(f"file://{os.path.realpath(index_file)}")
    else:
        logging.info("Index file not found.")

class Stream(QObject):
    newText = pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))

class WorkerThread(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, template_file, export_location, open_in_browser, use_whisper, slide_duration):
        super().__init__()
        self.template_file = os.path.normpath(template_file)
        self.export_location = os.path.normpath(export_location)
        self.open_in_browser = open_in_browser
        self.use_whisper = use_whisper
        self.slide_duration = slide_duration

    def run(self):
        try:
            self.progress.emit(5)
            self.status.emit("Initializing...")
            
            self.progress.emit(10)
            self.status.emit("Clearing export directory...")
            self.clear_directory(self.export_location)
            
            self.progress.emit(15)
            self.status.emit("Closing any open PowerPoint instances...")
            ensure_powerpoint_closed()
            
            self.progress.emit(20)
            self.status.emit("Exporting slides...")
            logging.info(f"Starting conversion with export location: {self.export_location}")
            transcribe_func = self.transcribe_audio if self.use_whisper else None
            
            slide_list = export_as_jpg_impl(self.template_file, self.export_location, self.slide_duration, self.progress.emit, transcribe_func)
            #slide_list = export_as_jpg_impl(self.template_file, self.export_location, default_slider_value, self.progress.emit, self.transcribe_audio if self.use_whisper else None)
            if slide_list:
                self.progress.emit(70)
                self.status.emit("Creating HTML...")
                create_html(self.export_location, slide_list, self.use_whisper)
                        
                if self.use_whisper:
                    self.progress.emit(75)
                    self.status.emit("Processing Whisper transcriptions...")
                    self.process_whisper_transcriptions(slide_list)
                
                self.progress.emit(90)
                self.status.emit("Finalizing...")
                
                logging.info(f"Files exported to: {self.export_location}")
                
                index_file = os.path.join(self.export_location, "index.html")
                if os.path.exists(index_file):
                    if self.open_in_browser:
                        self.status.emit("Opening in browser...")
                        open_html_in_browser(self.export_location)
                    else:
                        logging.info("HTML file created but not opened (as per user setting)")
                else:
                    logging.warning("index.html not found in export location.")
                
                self.progress.emit(100)
                self.status.emit("Conversion completed successfully!")
                self.finished.emit()
            else:
                self.error.emit("Failed to export slides.")
        except Exception as e:
            logging.error(f"Error in WorkerThread: {e}", exc_info=True)
            self.error.emit(str(e))

    def clear_directory(self, directory):
        if os.path.exists(directory):
            for filename in os.listdir(directory):
                file_path = os.path.join(directory, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    logging.info(f'Failed to delete {file_path}. Reason: {e}')
        else:
            os.makedirs(directory)

    def process_whisper_transcriptions(self, slide_list):
        for slide in slide_list:
            if slide['audio_path']:
                audio_path = os.path.join(self.export_location, slide['audio_path'])
                wav_path = self.convert_to_wav(audio_path)
                if wav_path:
                    transcription = self.transcribe_audio(wav_path)
                    slide['transcription'] = transcription

    def convert_to_wav(self, audio_path):
        if not os.path.exists(audio_path):
            logging.info(f"Audio file not found: {audio_path}")
            return None
        
        if audio_path.lower().endswith('.wav'):
            return audio_path
        
        wav_path = os.path.splitext(audio_path)[0] + '.wav'
        if os.path.exists(wav_path):
            return wav_path
        
        logging.info(f"Converting {audio_path} to {wav_path}")
        
        command = [
            ffmpeg_path,
            "-i", audio_path,
            "-acodec", "pcm_s16le",
            "-ar", "16000",
            "-ac", "1",
            "-f", "wav",
            wav_path
        ]
        
        try:
            subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            logging.info(f"Conversion successful: {wav_path}")
        except subprocess.CalledProcessError as e:
            logging.info(f"FFmpeg conversion failed: {e}")
            logging.info(f"FFmpeg stderr: {e.stderr.decode('utf-8')}")
            return None
        
        if os.path.exists(wav_path):
            return wav_path
        else:
            logging.info(f"WAV file not found after conversion: {wav_path}")
            return None

    def transcribe_audio(self, audio_path):
        if not audio_path or not os.path.exists(audio_path):
            logging.info(f"Audio file not found for transcription: {audio_path}")
            return None

        txt_path = audio_path + ".txt"
        if os.path.exists(txt_path):
            with open(txt_path, 'r', encoding='utf-8') as f:
                return f.read().strip()

        command = [
            whisper_exe,
            "-m", model_path,
            "-f", audio_path,
            "--output-txt"
        ]
        
        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            logging.info(f"Whisper output: {result.stdout}")
            return result.stdout.strip()
        except subprocess.CalledProcessError as e:
            logging.info(f"Error transcribing audio: {e}")
            logging.info(f"Stderr: {e.stderr}")
            return None

class QTextEditLogger(QObject, logging.Handler):
    new_record = pyqtSignal(str)

    def __init__(self, parent):
        super().__init__(parent)
        super(logging.Handler).__init__()
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        try:
            msg = self.format(record)
            self.new_record.emit(msg)
        except Exception as e:
            logging.info(f"Error in log handler: {e}")

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'PowerPoint to HTML Converter - Developed by Trevor Haagsma (v2.2.3) - General'
        self.left = 100
        self.top = 100
        self.width = 600
        self.height = 400
        self.settings_file = "converter_settings.json"
        self.initUI()
        self.setupLogger()
        self.redirect_output()
        self.load_settings()  # Load settings after initializing UI


    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        layout = QVBoxLayout()

        # File selection section
        file_section = QGroupBox("File Selection")
        file_layout = QVBoxLayout()
        file_explanation = QLabel("Select the PowerPoint file you want to convert.")
        file_explanation.setProperty("class", "explanation")
        self.file_label = QLabel('No file selected')
        self.select_button = QPushButton('Select PowerPoint File')
        self.select_button.clicked.connect(self.select_file)
        file_layout.addWidget(file_explanation)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.select_button)
        file_section.setLayout(file_layout)

        # Export section
        export_section = QGroupBox("Export Settings")
        export_layout = QVBoxLayout()
        export_explanation = QLabel("Choose where to save the converted files.  <span style='color: red;'>WARNING: This will delete everything currently in that folder.</span>")
        export_explanation.setTextFormat(Qt.RichText)
        export_explanation.setProperty("class", "explanation")
        export_layout.addWidget(export_explanation)
        
        export_path_layout = QHBoxLayout()
        self.export_location = QLineEdit(r'C:\Exported_Powerpoint')
        self.export_button = QPushButton('Browse')
        self.export_button.clicked.connect(self.select_export_location)
        export_path_layout.addWidget(self.export_location)
        export_path_layout.addWidget(self.export_button)
        export_layout.addLayout(export_path_layout)
        
        # Slider
        slider_layout = QVBoxLayout()
        slider_explanation = QLabel("Adjust the duration each slide will be displayed, note movies will not be adjusted & audio will join across slides.")
        slider_explanation.setProperty("class", "explanation")
        slider_layout.addWidget(slider_explanation)
        
        slider_control_layout = QHBoxLayout()
        self.duration_slider = QSlider(Qt.Horizontal)
        self.duration_slider.setMinimum(5)
        self.duration_slider.setMaximum(20)
        self.duration_slider.setValue(5)
        self.duration_slider.setTickPosition(QSlider.TicksBelow)
        self.duration_slider.setTickInterval(1)
        self.duration_label = QLabel('Slide Duration: 5 seconds')
        self.duration_slider.valueChanged.connect(self.update_duration_label)
        slider_control_layout.addWidget(self.duration_label)
        slider_control_layout.addWidget(self.duration_slider)
        slider_layout.addLayout(slider_control_layout)
        
        export_layout.addLayout(slider_layout)
        export_section.setLayout(export_layout)

        # Action section
        action_section = QGroupBox("Actions")
        action_layout = QVBoxLayout()
        action_explanation = QLabel("Start the conversion process and view progress.")
        action_explanation.setProperty("class", "explanation")
        action_layout.addWidget(action_explanation)
        self.convert_button = QPushButton('Convert')
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setEnabled(False)
        self.progress_bar = QProgressBar()
        action_layout.addWidget(self.convert_button)
        action_layout.addWidget(self.progress_bar)
        action_section.setLayout(action_layout)

        # Options section
        options_section = QGroupBox("Options")
        options_layout = QVBoxLayout()
        options_explanation = QLabel("Additional settings for the conversion process.")
        options_explanation.setProperty("class", "explanation")
        options_layout.addWidget(options_explanation)
        self.open_browser_checkbox = QCheckBox('Open in browser after conversion')
        self.open_browser_checkbox.setChecked(True)
        self.whisper_checkbox = QCheckBox('Use Whisper for transcriptions')
        self.whisper_checkbox.setChecked(False)
        options_layout.addWidget(self.open_browser_checkbox)
        self.reboot_checkbox = QCheckBox('Reboot remote machine after conversion')
        self.reboot_checkbox.setChecked(False)
        options_layout.addWidget(self.reboot_checkbox)

        # Optionally, add a Load Settings button
        #self.load_settings_button = QPushButton('Load Settings')
        #self.load_settings_button.clicked.connect(self.load_settings)
        #options_layout.addWidget(self.load_settings_button)

        #options_section.setLayout(options_layout)
        
        # Check if Whisper exe exists
        if not check_whisper_exe():
            self.whisper_checkbox.setEnabled(False)
            self.whisper_checkbox.setToolTip("Whisper executable not found")
        options_layout.addWidget(self.whisper_checkbox)
        
        # Check if 7-Zip is installed
        if not check_7zip():
            seven_zip_note = QLabel("Note: 7-Zip not found. Some features may not work properly.")                
            seven_zip_note.setStyleSheet("color: red;")
            options_layout.addWidget(seven_zip_note)

        # Check if ffmpeg is installed
        #if not check_ffmpeg_path():
        #    ffmpeg_note = QLabel("Note: ffmpeg not found. Some features may not work properly.")                
        #    ffmpeg_note.setStyleSheet("color: red;")
        #    options_layout.addWidget(ffmpeg_note)

        options_section.setLayout(options_layout)
        # Add Save Settings button
        self.save_settings_button = QPushButton('Save Settings')
        self.save_settings_button.clicked.connect(self.save_settings)
        
        # Add Save Settings button to options layout
        options_layout.addWidget(self.save_settings_button)

        # Log section
        log_section = QGroupBox("Logs")
        log_layout = QVBoxLayout()
        log_explanation = QLabel("View and export conversion logs.")
        log_explanation.setProperty("class", "explanation")
        log_layout.addWidget(log_explanation)
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.export_logs_button = QPushButton('Export Logs')
        self.export_logs_button.clicked.connect(self.export_logs)
        log_layout.addWidget(self.log_output)
        log_layout.addWidget(self.export_logs_button)
        log_section.setLayout(log_layout)

        # Add all sections to main layout
        layout.addWidget(file_section)
        layout.addWidget(export_section)
        layout.addWidget(action_section)
        layout.addWidget(options_section)
        layout.addWidget(log_section)

        self.status_label = QLabel('Ready')
        action_layout.addWidget(self.status_label)

        self.setLayout(layout)

    def save_settings(self):
        settings = {
            "export_location": os.path.normpath(self.export_location.text()),
            "slide_duration": self.duration_slider.value(),
            "open_browser": self.open_browser_checkbox.isChecked(),
            "use_whisper": self.whisper_checkbox.isChecked(),
            "file_path": os.path.normpath(getattr(self, 'template_file', "")),
            "reboot_machine": self.reboot_checkbox.isChecked()
        }
        
        logging.info(f"Saving settings: {settings}")
        
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f)
        
        logging.info(f"Settings saved to: {self.settings_file}")
        QMessageBox.information(self, "Settings Saved", "Your settings have been saved.")

    def save_current_settings(self):
        settings = {
            "export_location": os.path.normpath(self.export_location.text()),
            "slide_duration": self.duration_slider.value(),
            "open_browser": self.open_browser_checkbox.isChecked(),
            "use_whisper": self.whisper_checkbox.isChecked(),
            "file_path": os.path.normpath(getattr(self, 'template_file', ""))
        }
        
        logging.info(f"Saving current settings: {settings}")
        
        with open(self.settings_file, 'w') as f:
            json.dump(settings, f)
        
        logging.info(f"Current settings saved to: {self.settings_file}")

    def select_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select PowerPoint File", "", "PowerPoint Files (*.pptx);;All Files (*)", options=options)
        if file_name:
            self.file_label.setText(file_name)
            self.template_file = file_name
            self.convert_button.setEnabled(True)
            logging.info(f"File selected: {file_name}")
            self.save_current_settings()  # Save settings immediately after file selection
        else:
            logging.info("No file selected")
    def load_settings(self):
        logging.info("Starting to load settings")
        try:
            if os.path.exists(self.settings_file):
                logging.info(f"Settings file found: {self.settings_file}")
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                logging.info(f"Settings loaded: {settings}")
                
                if hasattr(self, 'export_location'):
                    export_location = os.path.normpath(settings.get("export_location", ""))
                    self.export_location.setText(export_location)
                    logging.info(f"Export location set to: {export_location}")

                if hasattr(self, 'duration_slider'):
                    self.duration_slider.setValue(settings.get("slide_duration", 5))
                    logging.info(f"Slide duration set to: {settings.get('slide_duration', 5)}")
                
                if hasattr(self, 'open_browser_checkbox'):
                    self.open_browser_checkbox.setChecked(settings.get("open_browser", True))
                    logging.info(f"Open browser checkbox set to: {settings.get('open_browser', True)}")
                
                if hasattr(self, 'whisper_checkbox') and hasattr(self, 'whisper_status_label'):
                    if check_whisper_exe():
                        self.whisper_checkbox.setChecked(settings.get("use_whisper", False))
                        self.whisper_status_label.setText("")
                        logging.info(f"Whisper checkbox set to: {settings.get('use_whisper', False)}")
                    else:
                        self.whisper_checkbox.setEnabled(False)
                        self.whisper_status_label.setText("Whisper executable not found")
                        self.whisper_status_label.setStyleSheet("color: gray;")
                        logging.info("Whisper executable not found")

                if hasattr(self, 'reboot_checkbox'):
                        self.reboot_checkbox.setChecked(settings.get("reboot_machine", False))
                
                saved_file = settings.get("file_path", "")
                logging.info(f"Saved file path from settings: {saved_file}")
                if hasattr(self, 'file_label') and hasattr(self, 'convert_button'):
                    if saved_file:
                        # Try the path as-is
                        if os.path.exists(saved_file):
                            file_path = saved_file
                        else:
                            # If it doesn't exist, try converting slashes
                            alternate_path = saved_file.replace('/', '\\') if '/' in saved_file else saved_file.replace('\\', '/')
                            file_path = alternate_path if os.path.exists(alternate_path) else None
                                    
                        if file_path:
                            self.file_label.setText(file_path)
                            self.template_file = file_path
                            self.convert_button.setEnabled(True)
                            logging.info(f"File loaded successfully: {file_path}")
                        else:
                            self.file_label.setText("No file selected")
                            self.template_file = None
                            self.convert_button.setEnabled(False)
                            logging.info(f"No valid file found in settings. Original path: {saved_file}")
                    else:
                        self.file_label.setText("No file selected")
                        self.template_file = None
                        self.convert_button.setEnabled(False)
                        logging.info("No file path found in settings")
                else:
                    logging.info("Settings file not found")
        except Exception as e:
            logging.error(f"Error loading settings: {e}", exc_info=True)

    def redirect_output(self):
        sys.stdout = Stream(newText=self.onUpdateLog)
        sys.stderr = Stream(newText=self.onUpdateLog)

    def update_duration_label(self, value):
        self.duration_label.setText(f'Slide Duration: {value} seconds')

    def setupLogger(self):
        self.logTextBox = QTextEditLogger(self)
        self.logTextBox.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(self.logTextBox)
        logging.getLogger().setLevel(logging.DEBUG)
        self.logTextBox.new_record.connect(self.onUpdateLog)

    @pyqtSlot(str)
    def onUpdateLog(self, text):
        self.log_output.append(text)

    def select_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select PowerPoint File", "", "PowerPoint Files (*.pptx);;All Files (*)", options=options)
        if file_name:
            file_name = os.path.normpath(file_name)
            self.file_label.setText(file_name)
            self.template_file = file_name
            self.convert_button.setEnabled(True)
            logging.info(f"File selected: {file_name}")
            self.save_current_settings()
        else:
            logging.info("No file selected")

    def select_export_location(self):
        export_dir = QFileDialog.getExistingDirectory(self, "Select Export Directory")
        if export_dir:
            # Normalize the path to use the correct separator for the current operating system
            export_dir = os.path.normpath(export_dir)
            self.export_location.setText(export_dir)
            logging.info(f"Export location selected: {export_dir}")
            self.save_current_settings()

    def start_conversion(self):
        self.convert_button.setEnabled(False)
        self.select_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText('Starting conversion...')
        open_in_browser = self.open_browser_checkbox.isChecked()
        use_whisper = self.whisper_checkbox.isChecked()
        slide_duration = self.duration_slider.value()
        export_location = os.path.normpath(self.export_location.text())
        self.worker = WorkerThread(self.template_file, self.export_location.text(), open_in_browser, use_whisper, slide_duration)
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.update_status)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.worker.start()
        logging.info(f"Conversion started with export location: {export_location}")

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def update_status(self, status):
        self.status_label.setText(status)

    def conversion_finished(self):
        self.progress_bar.setValue(99)
        self.status_label.setText("Conversion completed successfully!")
        self.convert_button.setEnabled(True)
        self.select_button.setEnabled(True)
        logging.info("Conversion completed successfully")

        if self.reboot_checkbox.isChecked():
            export_path = os.path.normpath(self.export_location.text())
            reboot = True  # Since the checkbox is checked, we always want to reboot

            # Check custom mappings first
            if export_path in CUSTOM_MAPPINGS:
                machine = CUSTOM_MAPPINGS[export_path]['machine']
                if send_reset_command(export_path, machine, reboot=reboot):
                    QMessageBox.information(self, "Reboot Command Sent", f"A reboot command has been sent to {machine}.")
                else:
                    QMessageBox.warning(self, "Reboot Command Failed", f"Failed to send reboot command to {machine}.")
                return

            # If not in custom mappings, check EXPORT_LOCATIONS
            for campus, locations in EXPORT_LOCATIONS.items():
                for path, details in locations.items():
                    if os.path.normpath(path) == export_path:
                        machine = details[1]  # Assuming the machine name is the second item in the tuple
                        if machine != "localhost":
                            if send_reset_command(export_path, machine, reboot=reboot):
                                QMessageBox.information(self, "Reboot Command Sent", f"A reboot command has been sent to {machine}.")
                            else:
                                QMessageBox.warning(self, "Reboot Command Failed", f"Failed to send reboot command to {machine}.")
                        return

            logging.warning(f"No matching machine found for export path: {export_path}")
        elif self.open_browser_checkbox.isChecked():
            open_html_in_browser(self.export_location.text())

    def conversion_error(self, error_msg):
        self.progress_bar.setValue(0)
        self.status_label.setText("Conversion failed")
        self.convert_button.setEnabled(True)
        self.select_button.setEnabled(True)
        logging.error(f"Conversion failed: {error_msg}")
        QMessageBox.critical(self, "Error", f"An error occurred during conversion:\n{error_msg}")

    def export_logs(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Log File", "", "Text Files (*.txt);;All Files (*)", options=options)
        if file_name:
            with open(file_name, 'w') as f:
                f.write(self.log_output.toPlainText())

    def select_export_location(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Export Location")
        if folder:
            self.export_location.setText(folder)
def main():
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    app = QApplication([])
    ex = App()
    ex.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

#Source for handling video playback errors
#https://stackoverflow.com/questions/49930680/how-to-handle-uncaught-in-promise-domexception-play-failed-because-the-use?rq=2
