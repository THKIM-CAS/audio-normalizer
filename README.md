# PPTX Narration Sound Tuner

A Python command-line tool that tunes audio volume and removes noise from PowerPoint narrations using the LUFS (Loudness Units relative to Full Scale) standard.

> **Note:** This is a modern uv-based Python project. All dependencies are managed with [uv](https://github.com/astral-sh/uv).

## Features

- **LUFS Normalization**: Uses EBU R128 standard for consistent perceived loudness
- **Audio Denoising**: Optional noise reduction to remove background HVAC, room tone, and ambient noise
- **Format Support**: Handles WAV, MP3, M4A, WMA, AAC, FLAC, and OGG audio formats
- **Non-Destructive**: Creates a new PPTX file, preserving the original
- **Batch Processing**: Process all PPTX files in a directory at once
- **Automatic Conversion**: Seamlessly handles format conversion via FFmpeg
- **Detailed Logging**: Provides clear feedback on normalization progress and results

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
# The tune-sound command will be available in the virtual environment
```

To run the tool after `uv sync`:

```bash
# Activate the virtual environment
source .venv/bin/activate  # macOS/Linux
# or
.venv\Scripts\activate     # Windows

# Then use the command
tune-sound input.pptx output.pptx
```

Alternatively, run directly with uv:

```bash
uv run tune-sound input.pptx output.pptx
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

After running `uv sync`, you can use the tool in several ways:

### Single File Mode

With uv (recommended):
```bash
uv run tune-sound input.pptx output.pptx
```

Or activate the virtual environment first:
```bash
source .venv/bin/activate  # macOS/Linux
tune-sound input.pptx output.pptx
```

This normalizes all audio in `input.pptx` to -16 LUFS and saves to `output.pptx`.

### Batch Directory Mode

Process all PPTX files in a directory:

```bash
uv run tune-sound --input-dir ./input --output-dir ./output
```

This will:
- Find all `.pptx` files in the input directory
- Normalize the audio in each file
- Save the normalized files to the output directory with the same filenames
- Create the output directory if it doesn't exist

### Custom Target LUFS

Single file:
```bash
uv run tune-sound input.pptx output.pptx --target-lufs -14
```

Batch mode:
```bash
uv run tune-sound --input-dir ./input --output-dir ./output --target-lufs -14
```

### Denoising Examples

Apply noise reduction with default strength:
```bash
uv run tune-sound input.pptx output.pptx --denoise
```

Apply noise reduction with custom strength:
```bash
uv run tune-sound input.pptx output.pptx --denoise --denoise-strength 0.7
```

Combine denoising with custom LUFS target:
```bash
uv run tune-sound input.pptx output.pptx --denoise --target-lufs -14
```

Batch processing with denoising:
```bash
uv run tune-sound \
  --input-dir ./input \
  --output-dir ./output \
  --denoise \
  --denoise-strength 0.6
```

### Verbose Output

```bash
uv run tune-sound input.pptx output.pptx --verbose
```

Shows detailed processing information including:
- Audio file detection
- Sample rates and durations
- Measured loudness levels
- Applied gain adjustments

### Force Overwrite

```bash
uv run tune-sound input.pptx output.pptx --force
```

Overwrites the output file(s) without prompting if they already exist. Works in both single file and batch modes.

### Complete Examples

Single file with all options:
```bash
uv run tune-sound \
  presentation.pptx \
  presentation_normalized.pptx \
  --target-lufs -16 \
  --verbose \
  --force
```

Batch processing with options:
```bash
uv run tune-sound \
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

## Command-Line Options

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

## How It Works

1. **Extraction**: The PPTX file (which is a ZIP archive) is extracted to a temporary directory
2. **Discovery**: Audio files are located in the `ppt/media/` directory
3. **Denoising** (optional): Background noise is removed using spectral gating
4. **Measurement**: Each audio file's loudness is measured using the EBU R128 algorithm
5. **Normalization**: Audio is adjusted to match the target LUFS level
6. **Reconstruction**: The modified PPTX is reassembled with normalized audio
7. **Output**: A new PPTX file is created with consistent audio levels

## Supported Audio Formats

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

## Output Examples

### Successful Single File Normalization

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

### Successful Batch Directory Processing

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

## Troubleshooting

### FFmpeg Not Found

**Error:**
```
ERROR: FFmpeg is not installed or not accessible.
```

**Solution:**
Install FFmpeg following the instructions in the Installation section above.

### No Audio Files Found

**Warning:**
```
WARNING: No audio files found in PPTX.
```

**Explanation:**
The PPTX file doesn't contain any slide narrations or audio files. A copy of the original will be created.

### Audio Too Short

**Warning:**
```
WARNING: Audio too short (0.2s) for LUFS measurement, skipping media1.wav
```

**Explanation:**
LUFS measurement requires at least ~400ms of audio. Very short clips are skipped.

### Invalid PPTX File

**Error:**
```
ERROR: Not a valid PPTX file (not a ZIP archive): file.pptx
```

**Solution:**
Ensure the file is a valid PowerPoint PPTX file, not a PPT (older format) or corrupted file.

### Permission Denied

**Error:**
```
ERROR: [Errno 13] Permission denied: 'output.pptx'
```

**Solution:**
- Ensure you have write permissions in the output directory
- Close the output file if it's open in PowerPoint
- Use a different output location

## Technical Details

### Project Structure

This is a modern Python package using `uv` for dependency management with a flat structure:

```
pptx_sound_normalizer/
├── pyproject.toml          # Project metadata and dependencies
├── README.md               # Documentation
├── __init__.py             # Package exports
├── cli.py                  # CLI entry point
├── audio_normalizer.py     # LUFS normalization engine
├── pptx_handler.py         # PPTX extraction/reconstruction
└── utils.py                # Helper utilities
```

### Architecture

The tool is built with a modular architecture:

- **cli.py**: Command-line interface and orchestration
- **pptx_handler.py**: PPTX extraction and reconstruction
- **audio_normalizer.py**: LUFS measurement and normalization
- **utils.py**: Helper functions and utilities

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

1. **Lossy Format Quality**: Converting MP3/M4A/WMA to WAV and back may introduce minor quality degradation
2. **Minimum Duration**: Audio clips shorter than ~400ms cannot be measured with LUFS
3. **Silent Audio**: Completely silent audio files are skipped
4. **File Size**: Very large PPTX files require sufficient disk space for temporary extraction

## PowerPoint to Video Export (Windows)

This project includes a VBScript for batch exporting PowerPoint presentations to video files on Windows.

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
