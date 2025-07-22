import os
from dotenv import load_dotenv
from glob import glob
import pandas as pd
from openai import OpenAI
from youtube_transcript_api import YouTubeTranscriptApi

PRESENTATION_FOLDER = 'presentation'
xlsx_files = glob(os.path.join(PRESENTATION_FOLDER, '*.xlsx'))
if not xlsx_files:
    raise FileNotFoundError(f"No .xlsx files found in '{PRESENTATION_FOLDER}'")
XLSX_FILE = xlsx_files[0]

load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY)

def time_to_seconds(t):
    parts = str(t).split(":")
    return int(parts[0]) * 60 + int(parts[1])

def summarize(text, prompt="Rewrite this text as narration:"):
    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "Create a summary for the following slide as if you were the one presenting. Get rid of any 'um's and 'ah's."},
                {"role": "user", "content": f"{prompt}\n\n{text}"}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"[OpenAI Error] {e}")
        return text

def generate_summaries(video_id, xlsx_path, output_folder="summaries"):
    os.makedirs(output_folder, exist_ok=True)

    # Load and prepare data
    df = pd.read_excel(xlsx_path)
    df['start_sec'] = df['Start time'].apply(time_to_seconds)

    # Get YouTube transcript
    transcript = YouTubeTranscriptApi.get_transcript(video_id)

    # Build time ranges for all slides
    slide_ranges = []
    for i in range(len(df)):
        start = df.loc[i, 'start_sec']
        if i + 1 < len(df):
            end = df.loc[i + 1, 'start_sec']
        else:
            end = float('inf')
        slide_ranges.append((start, end))

    # Process each slide
    for idx, (start, end) in enumerate(slide_ranges):
        slide_number = df.loc[idx, 'Slide Number']
        keep = df.loc[idx, 'skip/keep'] == 1

        lines = [entry['text'] for entry in transcript if start <= entry['start'] < end]
        raw_text = "\n".join(lines)

        if keep:
            summary = summarize(raw_text)
            summary_path = os.path.join(output_folder, f"slide_{slide_number}.txt")
            with open(summary_path, "w", encoding="utf-8") as f:
                f.write(summary)
            print(f"slide_{slide_number} summary saved.")
        else:
            print(f"Skipped slide_{slide_number}")

if __name__ == "__main__":
    video_id = input("Enter YouTube video ID: ").strip()
    generate_summaries(video_id, XLSX_FILE)
