# meetup-videos-remade
This project takes a PowerPoint presentation and a corresponding YouTube video, extracts the transcript, maps it to slides using an Excel timing file, generates narration using ElevenLabs, and creates a narrated video slideshow using FFmpeg


Instructions to Run:
1. In the working directory create a virtual enviorment and download requirements using 'pip install -r requirements.txt'
2. Install the following system tools: FFmpeg, LibreOffice, ImageMagick
    macOS:
        'brew install ffmpeg' 'brew install --cask LibreOffice' 'brew install imagemagick'
    Ubuntu/Debian/wsl:
        'sudo apt update'
        'sudo apt install ffmpeg libreoffice imagemagick'
3. Replace the text in '.env' with the actual values for the keys/ID
4. Place the presentation (.pptx) and the excel sheet with timestamps(.xlsx) in the presentation folder
5. run step 1 program using 'python step1.py'
6. paste the corresponding youtube video ID in the command line
example:
for the video https://www.youtube.com/live/wfbasy6k-CU
the ID would be: wfbasy6k-CU
7. Summarized text files will be placed in the summaries folder to edit
8. Run step 2 by using 'python step2.py'
9. output will be called 'final_output.mp4' and placed in the working directory

NOTE: To avoid using elevnlabs credits while testing comment out line 166 and 153. This will prevent elevenlabs from being used and from the separate slide images, audio files, and separated slide videos from being deleted.

If you want to save all individual outputs then comment out the last three lines of step2.py
