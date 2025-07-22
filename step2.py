import os
import subprocess
from dotenv import load_dotenv
from glob import glob
import requests
import pandas as pd
from openai import OpenAI

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
ELEVEN_API_KEY = os.getenv("ELEVEN_API_KEY")
ELEVEN_VOICE_ID = os.getenv("ELEVEN_VOICE_ID")
FFMPEG_PATH = os.getenv("FFMPEG_PATH", "ffmpeg")

client = OpenAI(api_key=OPENAI_API_KEY)

# Paths
PRESENTATION_FOLDER = 'presentation'
pptx_files = glob(os.path.join(PRESENTATION_FOLDER, '*.pptx'))
if not pptx_files:
    raise FileNotFoundError(f"No .pptx files found in '{PRESENTATION_FOLDER}'")
PPTX_FILE = pptx_files[0]

xlsx_files = glob(os.path.join(PRESENTATION_FOLDER, '*.xlsx'))
if not xlsx_files:
    raise FileNotFoundError(f"No .xlsx files found in '{PRESENTATION_FOLDER}'")
XLSX_FILE = xlsx_files[0]

# Read DataFrame to build mapping from slide rows to PNG filenames
df = pd.read_excel(XLSX_FILE)

df['png_name'] = df['Slide Number'].apply(lambda n: f"slide_{(n-1):02d}.png")

# Filter to only kept slides
keep_df = df[df['skip/keep'] == 1].copy()

TEXT_FOLDER = 'summaries'
WORK_DIR = 'output'
IMAGE_DIR = os.path.join(WORK_DIR, 'images')
pdf_filename = os.path.splitext(os.path.basename(PPTX_FILE))[0] + '.pdf'
PDF_PATH = os.path.join(WORK_DIR, pdf_filename)

os.makedirs(IMAGE_DIR, exist_ok=True)
os.makedirs(TEXT_FOLDER, exist_ok=True)


def convert_pptx_to_pdf_and_images(pptx_path, pdf_path, image_dir):
    print("Converting PPTX → PDF...")
    subprocess.run([
        'soffice', '--headless', '--convert-to', 'pdf', pptx_path,
        '--outdir', WORK_DIR
    ], check=True)

    print("Converting PDF → PNGs...")

    subprocess.run([
        'convert', '-density', '300', pdf_path,
        '-quality', '100', os.path.join(image_dir, 'slide_%02d.png')
    ], check=True)
    print(f"Slide images saved in: {image_dir}")


def generate_audio_for_slide(row):
    png_file = row['png_name']
    slide_idx = os.path.splitext(png_file)[0].split('_')[1]  
    text_path = os.path.join(TEXT_FOLDER, f"slide_{row['Slide Number']}.txt")
    audio_path = os.path.join(WORK_DIR, f"audio_{slide_idx}.mp3")

    if not os.path.exists(text_path):
        raise FileNotFoundError(f"Text not found for slide {row['Slide Number']}: {text_path}")

    with open(text_path, 'r') as sf:
        rewritten_text = sf.read()

    response = requests.post(
        f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVEN_VOICE_ID}",
        headers={"xi-api-key": ELEVEN_API_KEY, "Content-Type": "application/json"},
        json={
            "text": rewritten_text,
            "model_id": "eleven_multilingual_v2",
            "voice_settings": {"stability": 0.5, "similarity_boost": 0.75}
        }
    )
    if response.status_code != 200:
        raise Exception(f"Audio generation failed for slide {row['Slide Number']}: {response.text}")

    with open(audio_path, 'wb') as out:
        out.write(response.content)
    print(f"Audio saved: {audio_path}")


def generate_all_audio(keep_dataframe):
    for _, row in keep_dataframe.iterrows():
        print(f"Generating audio for slide {row['Slide Number']} (file {row['png_name']})...")
        generate_audio_for_slide(row)


def generate_video_for_slide(row):
    png_file = row['png_name']
    slide_idx = os.path.splitext(png_file)[0].split('_')[1]
    image_path = os.path.join(IMAGE_DIR, png_file)
    audio_path = os.path.join(WORK_DIR, f"audio_{slide_idx}.mp3")
    video_path = os.path.join(WORK_DIR, f"slide_{slide_idx}.mp4")

    if not os.path.exists(image_path) or not os.path.exists(audio_path):
        print(f"Skipping video for slide {row['Slide Number']}: missing {image_path if not os.path.exists(image_path) else audio_path}")
        return None

    print(f"Generating video for slide {row['Slide Number']}...")
    subprocess.run([
        FFMPEG_PATH, '-y',
        '-loop', '1', '-i', image_path,
        '-i', audio_path,
        '-c:v', 'libx264', '-tune', 'stillimage',
        '-c:a', 'aac', '-b:a', '192k',
        '-pix_fmt', 'yuv420p', '-shortest', video_path
    ], check=True)
    print(f"Video saved: {video_path}")
    return video_path


def generate_videos(keep_dataframe):
    videos = []
    for _, row in keep_dataframe.iterrows():
        vid = generate_video_for_slide(row)
        if vid:
            videos.append(vid)
    return videos


def concatenate_videos(video_paths, output_file):
    print("Concatenating videos...")
    concat_txt = os.path.join(WORK_DIR, 'concat_list.txt')
    with open(concat_txt, 'w') as f:
        for path in video_paths:
            f.write(f"file '{os.path.abspath(path)}'\n")
    cmd = [FFMPEG_PATH, '-f', 'concat', '-safe', '0', '-i', concat_txt, '-c', 'copy', output_file]
    subprocess.run(cmd, check=True)
    os.remove(concat_txt)
    print(f"Final video: {output_file}")


def clear_dir(folder):
    for fname in os.listdir(folder):
        path = os.path.join(folder, fname)
        if os.path.isfile(path):
            os.remove(path)


if __name__ == '__main__':
    convert_pptx_to_pdf_and_images(PPTX_FILE, PDF_PATH, IMAGE_DIR)
    generate_all_audio(keep_df)
    vids = generate_videos(keep_df)

    outro_files = glob(os.path.join('outro', '*.mp4'))
    if outro_files:
        outro = sorted(outro_files)[0]
        print(f"Appending outro video: {outro}")
        vids.append(outro)
    else:
        print("Warning: no outro found in 'outro/'")

    concatenate_videos(vids, 'final_output.mp4')
    clear_dir(TEXT_FOLDER)
    clear_dir(WORK_DIR)
    clear_dir('summaries')
