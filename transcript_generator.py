#!/usr/bin/env python3
"""
PPTX Narration Transcript Generator - Generate transcripts from PowerPoint narrations.

This tool extracts audio from PPTX files and generates text transcripts for each
slide's narration using OpenAI's Whisper speech-to-text model.
"""

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Dict, List, Optional

try:
    from faster_whisper import WhisperModel
    whisper_available = True
except ImportError:
    WhisperModel = None
    whisper_available = False

from pptx_handler import extract_pptx, find_audio_files
from utils import (
    TempDirectory,
    setup_logging,
    validate_pptx_file,
    find_pptx_files
)

logger = logging.getLogger('transcript_generator')


def transcribe_audio_file(audio_path: Path, model_name: str = "base",
                         model_cache: Dict = None) -> Optional[Dict]:
    """
    Transcribe an audio file using Whisper.

    Args:
        audio_path: Path to audio file
        model_name: Whisper model size (tiny, base, small, medium, large)
        model_cache: Optional dict to cache loaded models

    Returns:
        Dictionary with transcription results or None on error
    """
    if not whisper_available:
        logger.error("Whisper is not installed. Install with: uv pip install faster-whisper")
        return None

    try:
        # Load model (use cache to avoid reloading for batch processing)
        if model_cache is not None and model_name in model_cache:
            model = model_cache[model_name]
        else:
            logger.info(f"Loading Whisper model: {model_name}")
            model = WhisperModel(model_name, device="cpu", compute_type="int8")
            if model_cache is not None:
                model_cache[model_name] = model

        logger.info(f"Transcribing: {audio_path.name}")

        # Transcribe audio
        segments, info = model.transcribe(str(audio_path), beam_size=5)

        # Collect segments and build full text
        segment_list = []
        full_text = []

        for segment in segments:
            segment_list.append({
                'start': segment.start,
                'end': segment.end,
                'text': segment.text
            })
            full_text.append(segment.text)

        return {
            'text': ' '.join(full_text).strip(),
            'language': info.language,
            'segments': segment_list
        }
    except Exception as e:
        logger.error(f"Error transcribing {audio_path.name}: {e}")
        return None


def generate_transcript(pptx_path: Path, model_name: str = "base",
                       output_format: str = "txt") -> Optional[Dict[int, Dict]]:
    """
    Generate transcripts for all narrations in a PPTX file.

    Args:
        pptx_path: Path to PPTX file
        model_name: Whisper model size
        output_format: Output format (txt, json, or srt)

    Returns:
        Dictionary mapping slide numbers to transcription results
    """
    is_valid, error_msg = validate_pptx_file(pptx_path)
    if not is_valid:
        logger.error(error_msg)
        return None

    transcripts = {}
    model_cache = {}  # Cache model to avoid reloading

    try:
        with TempDirectory() as temp_dir:
            # Extract PPTX
            extract_path, audio_files = extract_pptx(pptx_path, temp_dir)

            if not audio_files:
                logger.warning("No audio files found in PPTX.")
                return {}

            logger.info(f"Found {len(audio_files)} audio file(s)")

            # Transcribe each audio file
            for idx, audio_file in enumerate(audio_files, start=1):
                logger.info(f"\nProcessing slide {idx} audio: {audio_file.name}")

                result = transcribe_audio_file(audio_file, model_name, model_cache)

                if result:
                    transcripts[idx] = {
                        'filename': audio_file.name,
                        'text': result['text'],
                        'language': result['language'],
                        'segments': result['segments']
                    }
                    logger.info(f"Slide {idx} transcribed successfully")
                else:
                    logger.warning(f"Failed to transcribe slide {idx}")

            return transcripts

    except Exception as e:
        logger.error(f"Error processing PPTX: {e}")
        return None


def save_transcript_txt(transcripts: Dict[int, Dict], output_path: Path) -> None:
    """Save transcripts as plain text file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("PowerPoint Narration Transcript\n")
        f.write("=" * 50 + "\n\n")

        for slide_num in sorted(transcripts.keys()):
            data = transcripts[slide_num]
            f.write(f"Slide {slide_num}\n")
            f.write("-" * 50 + "\n")
            f.write(f"Audio File: {data['filename']}\n")
            f.write(f"Language: {data['language']}\n")
            f.write(f"\nTranscript:\n{data['text']}\n\n")
            f.write("=" * 50 + "\n\n")

    logger.info(f"Transcript saved to: {output_path}")


def save_transcript_json(transcripts: Dict[int, Dict], output_path: Path) -> None:
    """Save transcripts as JSON file."""
    output_data = {
        'slides': transcripts,
        'total_slides': len(transcripts)
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)

    logger.info(f"Transcript saved to: {output_path}")


def save_transcript_srt(transcripts: Dict[int, Dict], output_path: Path) -> None:
    """Save transcripts as SRT subtitle file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        subtitle_index = 1

        for slide_num in sorted(transcripts.keys()):
            data = transcripts[slide_num]

            # Write slide header as a subtitle
            f.write(f"{subtitle_index}\n")
            f.write(f"00:00:00,000 --> 00:00:05,000\n")
            f.write(f"[Slide {slide_num}]\n\n")
            subtitle_index += 1

            # Write segments if available
            if data.get('segments'):
                for segment in data['segments']:
                    start_time = format_srt_time(segment['start'])
                    end_time = format_srt_time(segment['end'])
                    text = segment['text'].strip()

                    f.write(f"{subtitle_index}\n")
                    f.write(f"{start_time} --> {end_time}\n")
                    f.write(f"{text}\n\n")
                    subtitle_index += 1
            else:
                # If no segments, write full text
                f.write(f"{subtitle_index}\n")
                f.write(f"00:00:05,000 --> 00:00:10,000\n")
                f.write(f"{data['text']}\n\n")
                subtitle_index += 1

    logger.info(f"Transcript saved to: {output_path}")


def format_srt_time(seconds: float) -> str:
    """Convert seconds to SRT time format (HH:MM:SS,mmm)."""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    millis = int((seconds % 1) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"


def process_single_pptx(input_path: Path, output_path: Path, model_name: str,
                       output_format: str) -> int:
    """
    Process a single PPTX file to generate transcript.

    Args:
        input_path: Path to input PPTX file
        output_path: Path to output transcript file
        model_name: Whisper model size
        output_format: Output format (txt, json, or srt)

    Returns:
        0 on success, 1 on error
    """
    logger.info(f"Processing: {input_path}")

    transcripts = generate_transcript(input_path, model_name, output_format)

    if transcripts is None:
        return 1

    if not transcripts:
        logger.warning("No transcripts generated (no audio found).")
        return 0

    # Save transcript in requested format
    if output_format == "json":
        save_transcript_json(transcripts, output_path)
    elif output_format == "srt":
        save_transcript_srt(transcripts, output_path)
    else:  # txt
        save_transcript_txt(transcripts, output_path)

    logger.info(f"\nSuccessfully transcribed {len(transcripts)} slide(s)")
    return 0


def process_batch_pptx(input_dir: Path, output_dir: Path, model_name: str,
                      output_format: str) -> int:
    """
    Process all PPTX files in a directory.

    Args:
        input_dir: Directory containing PPTX files
        output_dir: Directory for output transcript files
        model_name: Whisper model size
        output_format: Output format (txt, json, or srt)

    Returns:
        0 on success, 1 on error
    """
    pptx_files = find_pptx_files(input_dir)

    if not pptx_files:
        logger.error(f"No PPTX files found in: {input_dir}")
        return 1

    logger.info(f"Found {len(pptx_files)} PPTX file(s) to process")

    output_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0

    for pptx_file in pptx_files:
        # Generate output filename
        ext_map = {"txt": ".txt", "json": ".json", "srt": ".srt"}
        output_file = output_dir / f"{pptx_file.stem}_transcript{ext_map[output_format]}"

        result = process_single_pptx(pptx_file, output_file, model_name, output_format)

        if result == 0:
            success_count += 1

        logger.info("-" * 60)

    logger.info(f"\nBatch processing complete: {success_count}/{len(pptx_files)} files processed")

    return 0 if success_count > 0 else 1


def main():
    """Main entry point for the transcript generator."""
    parser = argparse.ArgumentParser(
        description="Generate transcripts from PowerPoint narration audio using Whisper",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Generate transcript for a single PPTX file (text format)
  python transcript_generator.py presentation.pptx

  # Generate JSON format transcript
  python transcript_generator.py presentation.pptx -o transcript.json -f json

  # Generate SRT subtitle format
  python transcript_generator.py presentation.pptx -o transcript.srt -f srt

  # Use larger Whisper model for better accuracy
  python transcript_generator.py presentation.pptx -m medium

  # Process all PPTX files in a directory
  python transcript_generator.py input_folder/ -d output_transcripts/

Whisper Models (in order of size/accuracy):
  tiny   - Fastest, lowest accuracy (~1GB VRAM)
  base   - Default, good balance (~1GB VRAM)
  small  - Better accuracy (~2GB VRAM)
  medium - High accuracy (~5GB VRAM)
  large  - Best accuracy (~10GB VRAM)
        """
    )

    parser.add_argument(
        'input',
        type=Path,
        help='Input PPTX file or directory containing PPTX files'
    )

    parser.add_argument(
        '-o', '--output',
        type=Path,
        help='Output transcript file (for single file) or directory (for batch)'
    )

    parser.add_argument(
        '-d', '--output-dir',
        type=Path,
        help='Output directory for batch processing'
    )

    parser.add_argument(
        '-m', '--model',
        choices=['tiny', 'base', 'small', 'medium', 'large'],
        default='base',
        help='Whisper model size (default: base)'
    )

    parser.add_argument(
        '-f', '--format',
        choices=['txt', 'json', 'srt'],
        default='txt',
        help='Output format (default: txt)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )

    args = parser.parse_args()

    # Setup logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    setup_logging(log_level)

    # Check if Whisper is installed
    if not whisper_available:
        logger.error("faster-whisper is not installed.")
        logger.error("Install it with: uv sync")
        logger.error("Or: uv pip install faster-whisper")
        return 1

    # Determine if input is a file or directory
    if args.input.is_file():
        # Single file mode
        if args.output:
            output_path = args.output
        else:
            ext_map = {"txt": ".txt", "json": ".json", "srt": ".srt"}
            output_path = args.input.parent / f"{args.input.stem}_transcript{ext_map[args.format]}"

        return process_single_pptx(args.input, output_path, args.model, args.format)

    elif args.input.is_dir():
        # Batch mode
        output_dir = args.output_dir or args.output or (args.input / "transcripts")
        return process_batch_pptx(args.input, output_dir, args.model, args.format)

    else:
        logger.error(f"Input path does not exist: {args.input}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
