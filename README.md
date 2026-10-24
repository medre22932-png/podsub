# PodSub

PodSub is a desktop utility that processes subtitle files (`.srt`), plain text files (`.txt`), and Word documents (`.docx`), optimizing subtitle segmentation and converting them into clean, readable paragraph-formatted podcast transcripts.

## Features

- **Multi-Format Input**: Supports `.srt`, `.txt`, and `.docx` documents.
- **Tolerant Subtitle Parsing**: Interpolates broken or missing timestamps and handles custom timing formats.
- **Arabic Script Formatting**: 
  - Groups Arabic subtitles naturally without fragmenting words.
  - Strips trailing punctuation attached to Arabic script.
  - Automatically trims sentence-ending harakat on longer passages while preserving short quotations.
- **Batch Processing**: Select and process multiple files at once using a graphical file chooser.
- **Automated Output**: Generates cleanly formatted `podcast_<name>.txt` files in the source directory.

## Releases & Downloads

Pre-built binaries for **Windows (`podsub.exe`)** and **macOS (`podsub-macOS.zip`)** are automatically compiled on each release via GitHub Actions.

Check the [Latest Release](https://github.com/medre22932-png/podsub/releases/latest) page to download.

## Running from Source

### Prerequisites
- Python 3.10+
- Dependencies:
  ```bash
  pip install -r requirements.txt
  ```

### Run
```bash
python podsub.py
```