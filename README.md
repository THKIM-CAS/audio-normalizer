# Audio Normalizer Toolkit

A Python command-line toolkit for normalizing and denoising audio in both PowerPoint presentations (PPTX) and video files (MP4) using the LUFS (Loudness Units relative to Full Scale) standard, plus AI-powered transcription of PPTX narrations.

> **Note:** This is a modern uv-based Python project. All dependencies are managed with [uv](https://github.com/astral-sh/uv).

## What's Included

This toolkit provides four complementary tools:

1. **tune-pptx-sound** - PPTX audio normalizer and denoiser
2. **tune-video-sound** - Video (MP4) audio normalizer and denoiser
3. **pptx-transcript** - Generate transcripts from PPTX narration audio using AI
4. **export_pptx_to_video.vbs** - Windows batch script to export PPTX files to video

## Common Features

Both `tune-pptx-sound` and `tune-video-sound` share these core features:

- **LUFS Normalization**: Uses EBU R128 standard for consistent perceived loudness
- **Audio Denoising**: Optional noise reduction to remove background HVAC, room tone, and ambient noise
- **Batch Processing**: Process all files in a directory at once
- **Automatic Conversion**: Seamlessly handles format conversion via FFmpeg
- **Non-Destructive**: Creates new output files, preserving originals
- **Detailed Logging**: Provides clear feedback on normalization progress and results

### tune-pptx-sound Specific Features
- Handles WAV, MP3, M4A, WMA, AAC, FLAC, and OGG audio formats embedded in PPTX
- Preserves all PPTX structure and relationships
- Process all narrations across multiple slides at once

### tune-video-sound Specific Features
- Supports MP4 video files with h264/h265 codecs
- Video stream copied without re-encoding (fast processing)
- Audio extracted, processed, and replaced seamlessly

## Why LUFS?

LUFS (Loudness Units relative to Full Scale) measures perceived loudness, not just peak volume. This ensures:

- Consistent volume perception across all slides
- No jarring volume changes between narrations
- Professional audio quality following broadcast standards

**Common LUFS targets:**
- `-23 LUFS`: Broadcast standard (EBU R128)
- `-16 LUFS`: General use, e-learning (recommended)
- `-14 LUFS`: Streaming platforms (Spotify, YouTube)

## Audio Denoising

This tool includes optional audio denoising to remove background noise from your recordings:

- **Background Noise Removal**: Removes HVAC noise, room tone, and ambient sounds
- **Adjustable Strength**: Control denoising aggressiveness with `--denoise-strength` (0.0-1.0)
- **Smart Processing**: Denoising is applied before loudness measurement for better results
- **Optional**: Disabled by default - only activated with `--denoise` flag

**When to use denoising:**
- Recordings made in noisy environments
- Presentations with noticeable background hum or hiss
- Audio with HVAC or computer fan noise
- When you want cleaner, more professional sound

**Recommended strength values:**
- `0.3`: Light denoising for subtle background noise
- `0.5`: Medium denoising for moderate noise (default)
- `0.7-0.8`: Heavy denoising for very noisy recordings

## Requirements

- Python 3.8 or higher
- [uv](https://github.com/astral-sh/uv) - Modern Python package manager (required for this project)
- FFmpeg (for compressed audio format support)

## Installation

This project uses `uv` for fast, modern Python package management.

### 1. Install uv (if not already installed)

```bash
# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# Or with pip
pip install uv
```

### 2. Install the Project

```bash
cd pptx_sound_normalizer

# Sync dependencies and install the project (recommended)
uv sync

# This creates a virtual environment and installs all dependencies
# The tune-pptx-sound, tune-video-sound, and pptx-transcript commands will be available
```

To run the tools after `uv sync`:

```bash
# Activate the virtual environment
source .venv/bin/activate  # macOS/Linux
# or
.venv\Scripts\activate     # Windows

# Then use the commands
tune-pptx-sound input.pptx output.pptx
tune-video-sound input.mp4 output.mp4
pptx-transcript presentation.pptx
```

Alternatively, run directly with uv:

```bash
uv run tune-pptx-sound input.pptx output.pptx
uv run tune-video-sound input.mp4 output.mp4
uv run pptx-transcript presentation.pptx
```

### 3. Install FFmpeg

FFmpeg is required for processing compressed audio formats (MP3, M4A, WMA, etc.).

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install ffmpeg
```

**macOS:**
```bash
brew install ffmpeg
```

**Windows:**
- Download from [ffmpeg.org](https://ffmpeg.org/download.html)
- Or use Chocolatey: `choco install ffmpeg`

**Verify installation:**
```bash
ffmpeg -version
```

## Usage

After running `uv sync`, all three commands are available: `tune-pptx-sound`, `tune-video-sound`, and `pptx-transcript`.

---

## tune-pptx-sound: PPTX Audio Normalizer

Normalize and denoise audio narrations embedded in PowerPoint presentations.

### Single File Mode

With uv (recommended):
```bash
uv run tune-pptx-sound input.pptx output.pptx
```

Or activate the virtual environment first:
```bash
source .venv/bin/activate  # macOS/Linux
tune-pptx-sound input.pptx output.pptx
```

This normalizes all audio in `input.pptx` to -16 LUFS and saves to `output.pptx`.

### Batch Directory Mode

Process all PPTX files in a directory:

```bash
uv run tune-pptx-sound --input-dir ./input --output-dir ./output
```

This will:
- Find all `.pptx` files in the input directory
- Normalize the audio in each file
- Save the normalized files to the output directory with the same filenames
- Create the output directory if it doesn't exist

### Custom Target LUFS

Single file:
```bash
uv run tune-pptx-sound input.pptx output.pptx --target-lufs -14
```

Batch mode:
```bash
uv run tune-pptx-sound --input-dir ./input --output-dir ./output --target-lufs -14
```

### Denoising Examples

Apply noise reduction with default strength:
```bash
uv run tune-pptx-sound input.pptx output.pptx --denoise
```

Apply noise reduction with custom strength:
```bash
uv run tune-pptx-sound input.pptx output.pptx --denoise --denoise-strength 0.7
```

Combine denoising with custom LUFS target:
```bash
uv run tune-pptx-sound input.pptx output.pptx --denoise --target-lufs -14
```

Batch processing with denoising:
```bash
uv run tune-pptx-sound \
  --input-dir ./input \
  --output-dir ./output \
  --denoise \
  --denoise-strength 0.6
```

### Verbose Output

```bash
uv run tune-pptx-sound input.pptx output.pptx --verbose
```

Shows detailed processing information including:
- Audio file detection
- Sample rates and durations
- Measured loudness levels
- Applied gain adjustments

### Force Overwrite

```bash
uv run tune-pptx-sound input.pptx output.pptx --force
```

Overwrites the output file(s) without prompting if they already exist. Works in both single file and batch modes.

### Complete Examples

Single file with all options:
```bash
uv run tune-pptx-sound \
  presentation.pptx \
  presentation_normalized.pptx \
  --target-lufs -16 \
  --verbose \
  --force
```

Batch processing with options:
```bash
uv run tune-pptx-sound \
  --input-dir ./presentations \
  --output-dir ./normalized \
  --target-lufs -14 \
  --force
```

### Using as a Python Module

You can also import and use the modules directly in your own Python code:

```python
from pathlib import Path
from audio_normalizer import normalize_audio_files
from pptx_handler import extract_pptx, reconstruct_pptx
from utils import TempDirectory

input_file = Path("presentation.pptx")
output_file = Path("presentation_normalized.pptx")

with TempDirectory() as temp_dir:
    extract_path, audio_files = extract_pptx(input_file, temp_dir)
    stats = normalize_audio_files(audio_files, target_lufs=-16.0)
    reconstruct_pptx(extract_path, output_file)

    for stat in stats:
        print(f"{stat.filename}: {stat.original_lufs:.1f} → {stat.target_lufs:.1f} LUFS")
```

### tune-pptx-sound Command-Line Options

```
positional arguments:
  input                 Input PPTX file (single file mode)
  output                Output PPTX file (single file mode)

optional arguments:
  -h, --help            Show help message and exit
  --input-dir DIR       Input directory containing PPTX files (batch mode)
  --output-dir DIR      Output directory for normalized PPTX files (batch mode)
  --target-lufs FLOAT   Target loudness in LUFS (default: -16.0)
  --denoise             Apply noise reduction before normalization
  --denoise-strength FLOAT  Denoising strength: 0.0 (none) to 1.0 (maximum) (default: 0.5)
  -v, --verbose         Enable verbose logging
  -f, --force           Overwrite output file(s) without prompting
```

### Usage Modes

**Single File Mode**: Provide `input` and `output` positional arguments
**Batch Mode**: Use `--input-dir` and `--output-dir` options together

These modes are mutually exclusive - you cannot mix them.

### How tune-pptx-sound Works

1. **Extraction**: The PPTX file (which is a ZIP archive) is extracted to a temporary directory
2. **Discovery**: Audio files are located in the `ppt/media/` directory
3. **Denoising** (optional): Background noise is removed using spectral gating
4. **Measurement**: Each audio file's loudness is measured using the EBU R128 algorithm
5. **Normalization**: Audio is adjusted to match the target LUFS level
6. **Reconstruction**: The modified PPTX is reassembled with normalized audio
7. **Output**: A new PPTX file is created with consistent audio levels

### Supported Audio Formats in PPTX

| Format | Extension | Notes |
|--------|-----------|-------|
| WAV    | `.wav`    | Native support, no conversion needed |
| FLAC   | `.flac`   | Native support, lossless |
| OGG    | `.ogg`    | Native support |
| MP3    | `.mp3`    | Requires FFmpeg, lossy format |
| M4A    | `.m4a`    | Requires FFmpeg, lossy format |
| WMA    | `.wma`    | Requires FFmpeg, lossy format |
| AAC    | `.aac`    | Requires FFmpeg, lossy format |

**Note:** Compressed formats (MP3, M4A, WMA, AAC) undergo conversion to WAV for processing, then back to the original format. Some quality loss may occur with lossy formats.

### tune-pptx-sound Output Examples

#### Successful Single File Normalization

```
============================================================
PPTX Audio Normalizer
============================================================
FFmpeg: Available
Input:  presentation.pptx
Output: presentation_normalized.pptx
Target: -16.0 LUFS

INFO: Extracting PPTX: presentation.pptx
INFO: Found 5 audio file(s)

INFO: Processing: media1.m4a
INFO: Normalized media1.m4a: -22.3 LUFS → -16.0 LUFS (+6.3 dB)
INFO: Processing: media2.m4a
INFO: Normalized media2.m4a: -18.5 LUFS → -16.0 LUFS (+2.5 dB)
INFO: Processing: media3.m4a
INFO: Normalized media3.m4a: -12.1 LUFS → -16.0 LUFS (-3.9 dB)
INFO: Processing: media4.m4a
INFO: Normalized media4.m4a: -20.7 LUFS → -16.0 LUFS (+4.7 dB)
INFO: Processing: media5.m4a
INFO: Normalized media5.m4a: -19.2 LUFS → -16.0 LUFS (+3.2 dB)

------------------------------------------------------------
Successfully normalized 5 of 5 audio file(s):

  media1.m4a: -22.3 LUFS → -16.0 LUFS (+6.3 dB)
  media2.m4a: -18.5 LUFS → -16.0 LUFS (+2.5 dB)
  media3.m4a: -12.1 LUFS → -16.0 LUFS (-3.9 dB)
  media4.m4a: -20.7 LUFS → -16.0 LUFS (+4.7 dB)
  media5.m4a: -19.2 LUFS → -16.0 LUFS (+3.2 dB)

------------------------------------------------------------
INFO: Reconstructing PPTX: presentation_normalized.pptx
INFO: PPTX reconstruction complete: presentation_normalized.pptx

============================================================
Success! Output written to: presentation_normalized.pptx
============================================================
```

#### Successful Batch Directory Processing

```
============================================================
PPTX Audio Normalizer
Batch Directory Mode
============================================================
FFmpeg: Available
Target: -16.0 LUFS

Input Directory:  ./presentations
Output Directory: ./normalized
Found 3 PPTX file(s) to process

============================================================
Processing file 1/3: lecture1.pptx
============================================================
INFO: Extracting PPTX: presentations/lecture1.pptx
INFO: Found 8 audio file(s)
INFO: Processing: media1.m4a
INFO: Normalized media1.m4a: -22.3 LUFS → -16.0 LUFS (+6.3 dB)
...
INFO: Normalized 8 of 8 audio file(s)
INFO: Reconstructing PPTX: normalized/lecture1.pptx
INFO: Success! Output written to: normalized/lecture1.pptx

============================================================
Processing file 2/3: lecture2.pptx
============================================================
...

============================================================
Processing file 3/3: lecture3.pptx
============================================================
...

============================================================
Batch Processing Complete
============================================================
Total files:      3
Successfully processed: 3

```

### Troubleshooting tune-pptx-sound

#### FFmpeg Not Found

**Error:**
```
ERROR: FFmpeg is not installed or not accessible.
```

**Solution:**
Install FFmpeg following the instructions in the Installation section above.

#### No Audio Files Found

**Warning:**
```
WARNING: No audio files found in PPTX.
```

**Explanation:**
The PPTX file doesn't contain any slide narrations or audio files. A copy of the original will be created.

#### Audio Too Short

**Warning:**
```
WARNING: Audio too short (0.2s) for LUFS measurement, skipping media1.wav
```

**Explanation:**
LUFS measurement requires at least ~400ms of audio. Very short clips are skipped.

#### Invalid PPTX File

**Error:**
```
ERROR: Not a valid PPTX file (not a ZIP archive): file.pptx
```

**Solution:**
Ensure the file is a valid PowerPoint PPTX file, not a PPT (older format) or corrupted file.

#### Permission Denied

**Error:**
```
ERROR: [Errno 13] Permission denied: 'output.pptx'
```

**Solution:**
- Ensure you have write permissions in the output directory
- Close the output file if it's open in PowerPoint
- Use a different output location

---

## pptx-transcript: Transcript Generator

Generate text transcripts from PowerPoint narration audio using OpenAI's Whisper AI model for speech-to-text conversion.

### Features

- **AI-Powered Transcription**: Uses OpenAI's Whisper model for accurate speech-to-text
- **Multiple Model Sizes**: Choose from tiny to large models based on accuracy/speed tradeoffs
- **Multiple Output Formats**: Generate transcripts in TXT, JSON, or SRT (subtitle) format
- **Slide-by-Slide Organization**: Transcripts are organized by slide number
- **Batch Processing**: Process all PPTX files in a directory at once
- **Language Detection**: Automatically detects the spoken language

### Basic Usage

Generate a transcript for a single PPTX file:

```bash
uv run pptx-transcript presentation.pptx
```

This creates `presentation_transcript.txt` with transcripts for all slide narrations.

Or activate the virtual environment first:
```bash
source .venv/bin/activate  # macOS/Linux
pptx-transcript presentation.pptx
```

### Output Formats

**Text format (default):**
```bash
uv run pptx-transcript presentation.pptx -f txt
```

**JSON format:**
```bash
uv run pptx-transcript presentation.pptx -f json -o transcript.json
```

**SRT subtitle format:**
```bash
uv run pptx-transcript presentation.pptx -f srt -o transcript.srt
```

### Whisper Model Selection

Choose a model based on your needs:

```bash
# Fast, less accurate (good for quick drafts)
uv run pptx-transcript presentation.pptx -m tiny

# Default - good balance of speed and accuracy
uv run pptx-transcript presentation.pptx -m base

# Better accuracy, slower
uv run pptx-transcript presentation.pptx -m small

# High accuracy
uv run pptx-transcript presentation.pptx -m medium

# Best accuracy, slowest
uv run pptx-transcript presentation.pptx -m large
```

**Model comparison:**
- `tiny` - Fastest, ~1GB VRAM, good for quick drafts
- `base` - Default, ~1GB VRAM, good balance
- `small` - ~2GB VRAM, better accuracy
- `medium` - ~5GB VRAM, high accuracy
- `large` - ~10GB VRAM, best accuracy

### Batch Processing

Process all PPTX files in a directory:

```bash
uv run pptx-transcript presentations/ -d transcripts/
```

This will generate transcripts for all PPTX files found in the `presentations/` directory and save them to `transcripts/`.

### Custom Output Path

Specify a custom output file:

```bash
uv run pptx-transcript presentation.pptx -o my_transcript.txt
```

### Verbose Output

Enable detailed logging:

```bash
uv run pptx-transcript presentation.pptx -v
```

### Complete Example

```bash
uv run pptx-transcript \
  presentation.pptx \
  -o presentation_transcript.json \
  -f json \
  -m medium \
  -v
```

### Output Format Examples

**TXT Format:**
```
PowerPoint Narration Transcript
==================================================

Slide 1
--------------------------------------------------
Audio File: media1.m4a
Language: english

Transcript:
Welcome to this presentation on audio normalization.

==================================================

Slide 2
--------------------------------------------------
Audio File: media2.m4a
Language: english

Transcript:
Today we'll cover LUFS standards and best practices.

==================================================
```

**JSON Format:**
```json
{
  "slides": {
    "1": {
      "filename": "media1.m4a",
      "text": "Welcome to this presentation on audio normalization.",
      "language": "english",
      "segments": [...]
    },
    "2": {
      "filename": "media2.m4a",
      "text": "Today we'll cover LUFS standards and best practices.",
      "language": "english",
      "segments": [...]
    }
  },
  "total_slides": 2
}
```

**SRT Format:**
```
1
00:00:00,000 --> 00:00:05,000
[Slide 1]

2
00:00:05,000 --> 00:00:10,500
Welcome to this presentation on audio normalization.

3
00:00:00,000 --> 00:00:05,000
[Slide 2]

4
00:00:05,000 --> 00:00:12,300
Today we'll cover LUFS standards and best practices.
```

### Command-Line Options

```
positional arguments:
  input                 Input PPTX file or directory containing PPTX files

optional arguments:
  -h, --help            Show help message and exit
  -o, --output PATH     Output transcript file (for single file) or directory (for batch)
  -d, --output-dir DIR  Output directory for batch processing
  -m, --model MODEL     Whisper model size: tiny, base, small, medium, large (default: base)
  -f, --format FORMAT   Output format: txt, json, srt (default: txt)
  -v, --verbose         Enable verbose logging
```

### How It Works

1. **Extraction**: PPTX file is extracted to access embedded audio files
2. **Audio Discovery**: Finds all narration audio files in the presentation
3. **AI Transcription**: Each audio file is transcribed using Whisper AI
4. **Organization**: Transcripts are organized by slide number
5. **Output**: Generates transcript file in the requested format

### Requirements

The transcript generator requires the `faster-whisper` package, which is included in the project dependencies. This is a faster, more efficient implementation of OpenAI's Whisper model that's fully compatible with Python 3.12. On first run, Whisper will download the selected model (file size varies by model).

---

## tune-video-sound: Video Audio Normalizer

Normalize and denoise audio tracks in MP4 video files. The video stream is copied without re-encoding for fast processing.

### Single File Mode

With uv (recommended):
```bash
uv run tune-video-sound input.mp4 output.mp4
```

Or activate the virtual environment first:
```bash
source .venv/bin/activate  # macOS/Linux
tune-video-sound input.mp4 output.mp4
```

This extracts audio from `input.mp4`, normalizes it to -16 LUFS, and saves to `output.mp4`.

### Batch Directory Mode

Process all MP4 files in a directory:

```bash
uv run tune-video-sound --input-dir ./videos --output-dir ./processed
```

This will:
- Find all `.mp4` files in the input directory
- Extract and normalize the audio in each file
- Save the processed files to the output directory with the same filenames
- Create the output directory if it doesn't exist

### Custom Target LUFS

Single file:
```bash
uv run tune-video-sound input.mp4 output.mp4 --target-lufs -14
```

Batch mode:
```bash
uv run tune-video-sound --input-dir ./videos --output-dir ./processed --target-lufs -14
```

### Denoising Examples

Apply noise reduction with default strength:
```bash
uv run tune-video-sound input.mp4 output.mp4 --denoise
```

Apply noise reduction with custom strength:
```bash
uv run tune-video-sound input.mp4 output.mp4 --denoise --denoise-strength 0.7
```

Combine denoising with custom LUFS target:
```bash
uv run tune-video-sound input.mp4 output.mp4 --denoise --target-lufs -14
```

Batch processing with denoising:
```bash
uv run tune-video-sound \
  --input-dir ./videos \
  --output-dir ./processed \
  --denoise \
  --denoise-strength 0.6
```

### Verbose Output

```bash
uv run tune-video-sound input.mp4 output.mp4 --verbose
```

Shows detailed processing information including:
- Audio extraction details
- Sample rates and durations
- Measured loudness levels
- Applied gain adjustments

### Force Overwrite

```bash
uv run tune-video-sound input.mp4 output.mp4 --force
```

Overwrites the output file(s) without prompting if they already exist. Works in both single file and batch modes.

### Complete Examples

Single file with all options:
```bash
uv run tune-video-sound \
  lecture.mp4 \
  lecture_normalized.mp4 \
  --target-lufs -16 \
  --denoise \
  --denoise-strength 0.5 \
  --verbose \
  --force
```

Batch processing with options:
```bash
uv run tune-video-sound \
  --input-dir ./recordings \
  --output-dir ./normalized \
  --target-lufs -14 \
  --denoise \
  --force
```

### tune-video-sound Command-Line Options

```
positional arguments:
  input                 Input MP4 file (single file mode)
  output                Output MP4 file (single file mode)

optional arguments:
  -h, --help            Show help message and exit
  --input-dir DIR       Input directory containing MP4 files (batch mode)
  --output-dir DIR      Output directory for processed videos (batch mode)
  --target-lufs FLOAT   Target loudness in LUFS (default: -16.0)
  --denoise             Apply noise reduction before normalization
  --denoise-strength FLOAT  Denoising strength: 0.0 (none) to 1.0 (maximum) (default: 0.5)
  -v, --verbose         Enable verbose logging
  -f, --force           Overwrite output file(s) without prompting
```

### Usage Modes

**Single File Mode**: Provide `input` and `output` positional arguments
**Batch Mode**: Use `--input-dir` and `--output-dir` options together

These modes are mutually exclusive - you cannot mix them.

### How tune-video-sound Works

1. **Validation**: Checks that the video file exists and contains both video and audio streams
2. **Audio Extraction**: Extracts audio track to temporary WAV file using FFmpeg (48kHz stereo)
3. **Denoising** (optional): Background noise is removed using spectral gating
4. **Measurement**: Audio loudness is measured using the EBU R128 algorithm
5. **Normalization**: Audio is adjusted to match the target LUFS level
6. **Audio Replacement**: Normalized audio replaces original track (video stream copied, no re-encoding)
7. **Output**: A new MP4 file is created with normalized audio and original video

### Supported Video Formats

| Format | Extension | Video Codec | Notes |
|--------|-----------|-------------|-------|
| MP4    | `.mp4`    | h264, h265  | Video copied without re-encoding; audio replaced with AAC 192kbps |

**Note:** The video stream is copied directly (no re-encoding) for fast processing. Only the audio track is processed.

### tune-video-sound Output Examples

#### Successful Single File Processing

```
============================================================
Video Audio Normalizer
============================================================
FFmpeg: Available
Target: -16.0 LUFS
Denoise: Enabled (strength: 0.50)

Input:  lecture_recording.mp4
Output: lecture_normalized.mp4

INFO: Extracting audio from: lecture_recording.mp4

INFO: Processing audio...
INFO: Measured loudness: -22.3 LUFS
INFO: Normalizing to -16.0 LUFS (gain: +6.3 dB)
INFO: Audio normalization complete

INFO: Replacing audio in: lecture_recording.mp4
INFO: Success! Output written to: lecture_normalized.mp4

------------------------------------------------------------
Normalization: -22.3 LUFS → -16.0 LUFS (+6.3 dB)
------------------------------------------------------------

============================================================
Processing Complete
============================================================
```

#### Successful Batch Directory Processing

```
============================================================
Video Audio Normalizer
Batch Directory Mode
============================================================
FFmpeg: Available
Target: -16.0 LUFS
Denoise: Disabled

Input Directory:  ./recordings
Output Directory: ./processed
Found 3 video file(s) to process

============================================================
Processing file 1/3: lecture1.mp4
============================================================
INFO: Extracting audio from: lecture1.mp4
INFO: Measured loudness: -20.1 LUFS
INFO: Normalizing to -16.0 LUFS (gain: +4.1 dB)
INFO: Success! Output written to: processed/lecture1.mp4

============================================================
Processing file 2/3: lecture2.mp4
============================================================
...

============================================================
Processing file 3/3: lecture3.mp4
============================================================
...

============================================================
Batch Processing Complete
============================================================
Total files:            3
Successfully processed: 3
```

### Troubleshooting tune-video-sound

#### FFmpeg Not Found

**Error:**
```
ERROR: FFmpeg is not installed or not accessible.
```

**Solution:**
Install FFmpeg following the instructions in the Installation section above. FFmpeg is required for video/audio extraction and replacement.

#### No Audio Stream

**Error:**
```
ERROR: No audio stream found in: video.mp4
```

**Explanation:**
The video file doesn't contain an audio track. tune-video requires videos with audio.

**Solution:**
- Verify the video file has audio using a media player
- Check if the audio track was accidentally removed during editing
- Use a different video file

#### No Video Stream

**Error:**
```
ERROR: No video stream found in: file.mp4
```

**Explanation:**
The file doesn't contain a video stream (might be audio-only).

**Solution:**
- Verify this is a valid video file
- For audio-only files, consider using a different tool or converting to audio format first

#### Unsupported Format

**Error:**
```
ERROR: Unsupported format: .avi (only .mp4 supported)
```

**Solution:**
Currently only MP4 format is supported. Convert your video to MP4 format first using FFmpeg:
```bash
ffmpeg -i input.avi -c:v libx264 -c:a aac output.mp4
```

#### Permission Denied

**Error:**
```
ERROR: [Errno 13] Permission denied: 'output.mp4'
```

**Solution:**
- Ensure you have write permissions in the output directory
- Close the output file if it's open in a video player
- Use a different output location

---

## Technical Details

### Project Structure

This is a modern Python package using `uv` for dependency management with a flat structure:

```
audio-normalizer-toolkit/
├── pyproject.toml          # Project metadata and dependencies
├── README.md               # Documentation
├── __init__.py             # Package exports
├── cli.py                  # tune-pptx-sound CLI entry point
├── video_cli.py            # tune-video-sound CLI entry point
├── audio_normalizer.py     # LUFS normalization engine
├── pptx_handler.py         # PPTX extraction/reconstruction
├── video_handler.py        # Video audio extraction/replacement
├── utils.py                # Helper utilities
├── export_pptx_to_video.vbs # Windows PPTX to video export
└── run_export.bat          # Windows batch wrapper
```

### Architecture

The toolkit is built with a modular architecture shared between both tools:

**Shared modules:**
- **audio_normalizer.py**: LUFS measurement and normalization (used by both tools)
- **utils.py**: Helper functions and utilities (used by both tools)

**tune-pptx-sound specific:**
- **cli.py**: Command-line interface for PPTX processing
- **pptx_handler.py**: PPTX extraction and reconstruction

**tune-video-sound specific:**
- **video_cli.py**: Command-line interface for video processing
- **video_handler.py**: Video audio extraction and replacement via FFmpeg

### LUFS Normalization Algorithm

The tool uses `pyloudnorm`, a Python implementation of ITU-R BS.1770-4:

1. Load audio data
2. Measure integrated loudness (LUFS) using a calibrated loudness meter
3. Calculate required gain: `gain = target_lufs - measured_lufs`
4. Apply gain while preventing clipping (true peak limiting at -1.0 dBTP)
5. Write normalized audio

### PPTX Structure Preservation

- PPTX files are ZIP archives with a specific structure
- Audio files are stored in `ppt/media/` directory
- XML relationship files map media to slides
- The tool preserves all file names and structure to maintain relationships
- Only audio data is modified; all other content remains unchanged

## Limitations

### tune-pptx-sound Limitations
1. **Lossy Format Quality**: Converting MP3/M4A/WMA to WAV and back may introduce minor quality degradation
2. **Minimum Duration**: Audio clips shorter than ~400ms cannot be measured with LUFS
3. **Silent Audio**: Completely silent audio files are skipped
4. **File Size**: Very large PPTX files require sufficient disk space for temporary extraction

### tune-video-sound Limitations
1. **Format Support**: Currently only MP4 files are supported (h264/h265 codecs)
2. **Audio Re-encoding**: Audio is re-encoded to AAC 192kbps (minor quality change possible)
3. **Minimum Duration**: Audio clips shorter than ~400ms cannot be measured with LUFS
4. **Silent Video**: Videos with completely silent audio are skipped

---

## Bonus: PowerPoint to Video Export (Windows)

This toolkit includes a Windows VBScript for batch exporting PowerPoint presentations to video files, which complements the audio normalization workflow.

### Typical Workflow

1. **Normalize PPTX Audio**: Use `tune-pptx-sound` to normalize audio in your PowerPoint presentations
2. **Export to Video**: Use `export_pptx_to_video.vbs` to convert PPTX files to MP4 videos
3. **Fine-tune Video Audio** (optional): Use `tune-video-sound` if additional audio adjustments are needed

### Files

- **export_pptx_to_video.vbs** - Main VBScript for batch export
- **run_export.bat** - Convenient batch file wrapper

### Features

- Batch processes all .pptx and .ppt files in a directory
- Exports to MP4 format with configurable quality (HD 720p default)
- Supports slide timings or default duration per slide
- Progress tracking with detailed status reporting
- Separate input/output directory support

### Requirements

- Windows operating system
- Microsoft PowerPoint installed
- PowerPoint COM automation support

### Usage

**Option 1: Double-click with GUI prompts**
```
Double-click run_export.bat and enter paths when prompted
```

**Option 2: Command line with arguments**
```batch
run_export.bat "C:\path\to\pptx\files" "C:\path\to\output"
```

**Option 3: Using VBScript directly**
```batch
cscript export_pptx_to_video.vbs "C:\path\to\pptx\files" "C:\path\to\output"
```

### Configuration

To change video quality, edit line 17 in `export_pptx_to_video.vbs`:

- `videoQuality = 1` - Low (480p, 24fps)
- `videoQuality = 2` - Medium (576p, 25fps)
- `videoQuality = 3` - High (720p, 25fps)
- `videoQuality = 4` - Very High (720p, 30fps)
- `videoQuality = 5` - HD 720p (default, 30fps)
- `videoQuality = 6` - Full HD 1080p (30fps)

### Output

Creates .mp4 video files with the same base names as the original PowerPoint files. The script displays real-time progress and provides a summary of successful and failed exports.

## Contributing

Contributions are welcome! Feel free to:

- Report bugs or issues
- Suggest new features
- Submit pull requests

## License

This project is provided as-is for educational and professional use.

## Acknowledgments

- [pyloudnorm](https://github.com/csteinmetz1/pyloudnorm): EBU R128 loudness normalization
- [soundfile](https://python-soundfile.readthedocs.io/): Audio I/O library
- [FFmpeg](https://ffmpeg.org/): Audio format conversion

## Contact

For questions, issues, or feedback, please open an issue in the repository.
