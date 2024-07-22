# PPTX2HTML
Powerpoint to HTML converter

For the experimental Whisper transcriptions to work, you need to have the following files in the same directory (note whisper is in a folder called 'Whisper'): 
"ffmpeg.exe"
"Whisper\\main.exe"
"Whisper\\ggml-model-whisper-base.bin"
This is set to use the lite model. 
7zip is only required if a CRC error is encountered as i couldnt utilise the zipfile package to handle CRC validation error. This script assumed your install path for 7zip is C:\\Program Files\\7-Zip\\7z.exe

Adjusting the duration of the slides will not impact the movie size, so if the movie is 1 hour, the slideshow will not move to the next html file until that movie has finished. 
Audio joins across slides with a slight delay. 
Work needs to be done to get the transcriptions lined up completely.

![image](https://github.com/user-attachments/assets/2c8232ff-d421-42e9-b0b3-cde563251f24)

