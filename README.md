# The Crucible (Archive)

**Originally built:** Fall 2024 • **Status:** Archived • **Stack:** Python, SQLite, OpenPyXL, FFmpeg, PIL

## What it does
**The Crucible** automates the reconciliation of Baselight exports and Xytech work orders.  
It validates frame ranges, matches editorial and production data, and generates  
Excel/CSV reports complete with **timestamps and thumbnails** pulled from the source video.  

All that’s required is a **Baselight export**, a **Xytech text file**, and a **video file**.  
This project was built for **post-production workflows** where data from multiple  
systems must be synchronized and validated quickly.  
Features

Parse Baselight data – reels, parts, frame ranges.

Parse Xytech data – producer, operator, job metadata.

SQLite integration – store matched datasets.

Frame validation – against video duration (fps = 24).

Thumbnail generation – via ffmpeg, embedded in Excel.

Unused frame reporting – flagged and exported to CSV.

Excel export – producer, operator, order info, frame ranges, timecodes, thumbnails.

File overview

thecrucible.py – main processing pipeline.

Baselight_export_fall2024.txt – sample Baselight export.

Xytech_fall2024.txt – sample Xytech order.

thecrucible_database.db – SQLite database output.

combined_output_with_images.xlsx – Excel output (thumbnails + metadata).

## Quickstart
```bash
# prerequisites
Python 3.10+
pip install -r requirements.txt

# run
python thecrucible.py \
  --baselight Baselight_export_fall2024.txt \
  --xytech Xytech_fall2024.txt \
  --process FROM-ZERO_Demo-REEL_V01-04.mp4 \
  --outputXLS combined_output_with_images.xlsx \
  --outputCSV unused_data.csv
