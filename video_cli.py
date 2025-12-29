#!/usr/bin/env python3
"""
Video Audio Normalizer - Normalize and denoise audio in video files.

This tool extracts audio from video files, optionally removes background noise,
normalizes it to a target LUFS (Loudness Units relative to Full Scale) level,
and replaces the audio track back into the video.
"""

import argparse
import sys
from pathlib import Path

from audio_normalizer import normalize_audio_file
from video_handler import (
    extract_audio_from_video,
    replace_audio_in_video,
    validate_video_file,
    find_video_files
)
from utils import (
    TempDirectory,
    setup_logging,
    check_ffmpeg_installed,
    confirm_overwrite,
    format_lufs
)


def process_single_video(input_path: Path, output_path: Path, target_lufs: float,
                        force: bool, denoise: bool, denoise_strength: float,
                        logger) -> int:
    """
    Process a single video file.

    Args:
        input_path: Path to input video file
        output_path: Path to output video file
        target_lufs: Target loudness in LUFS
        force: Overwrite output without prompting
        denoise: Whether to apply denoising
        denoise_strength: Denoising strength (0.0-1.0)
        logger: Logger instance

    Returns:
        0 on success, 1 on error
    """
    # Validate input file
    is_valid, error_msg = validate_video_file(input_path)
    if not is_valid:
        logger.error(error_msg)
        return 1

    # Check if output file exists
    if output_path.exists() and not force:
        if not confirm_overwrite(output_path):
            logger.info("Skipped (file exists).")
            return 0

    try:
        with TempDirectory() as temp_dir:
            # Extract audio from video
            temp_audio = temp_dir / "extracted_audio.wav"
            extract_audio_from_video(input_path, temp_audio)

            logger.info("")

            # Normalize audio file
            stats = normalize_audio_file(temp_audio, target_lufs, denoise, denoise_strength)

            if not stats:
                logger.warning("Audio normalization failed or was skipped.")
                return 1

            logger.info("")

            # Replace audio in video
            replace_audio_in_video(input_path, temp_audio, output_path)
            logger.info(f"Success! Output written to: {output_path}")

            logger.info("")
            logger.info("-" * 60)
            logger.info(f"Normalization: {stats}")
            logger.info("-" * 60)

            return 0

    except Exception as e:
        logger.error(f"Failed to process {input_path.name}: {e}")
        return 1


def process_batch_directory(input_dir: Path, output_dir: Path, target_lufs: float,
                            force: bool, denoise: bool, denoise_strength: float,
                            logger) -> int:
    """
    Process all video files in a directory.

    Args:
        input_dir: Input directory containing video files
        output_dir: Output directory for processed videos
        target_lufs: Target loudness in LUFS
        force: Overwrite output files without prompting
        denoise: Whether to apply denoising
        denoise_strength: Denoising strength (0.0-1.0)
        logger: Logger instance

    Returns:
        0 on success, 1 on error
    """
    # Validate input directory
    if not input_dir.exists():
        logger.error(f"Input directory not found: {input_dir}")
        return 1

    if not input_dir.is_dir():
        logger.error(f"Input path is not a directory: {input_dir}")
        return 1

    # Find all video files
    video_files = find_video_files(input_dir)

    if not video_files:
        logger.warning(f"No MP4 files found in: {input_dir}")
        return 0

    logger.info(f"Input Directory:  {input_dir}")
    logger.info(f"Output Directory: {output_dir}")
    logger.info(f"Found {len(video_files)} video file(s) to process")
    logger.info("")

    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Process each video file
    success_count = 0
    error_count = 0

    for i, input_file in enumerate(video_files, 1):
        logger.info("=" * 60)
        logger.info(f"Processing file {i}/{len(video_files)}: {input_file.name}")
        logger.info("=" * 60)

        output_file = output_dir / input_file.name

        result = process_single_video(input_file, output_file, target_lufs, force,
                                     denoise, denoise_strength, logger)

        if result == 0:
            success_count += 1
        else:
            error_count += 1

        logger.info("")

    # Final summary
    logger.info("=" * 60)
    logger.info("Batch Processing Complete")
    logger.info("=" * 60)
    logger.info(f"Total files:            {len(video_files)}")
    logger.info(f"Successfully processed: {success_count}")
    if error_count > 0:
        logger.info(f"Errors:                 {error_count}")
    logger.info("")

    return 0 if error_count == 0 else 1


def main():
    """Main entry point for the video audio normalizer."""
    parser = argparse.ArgumentParser(
        description='Normalize and denoise audio in video files (MP4) using LUFS standard.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Single file mode:
    %(prog)s input.mp4 output.mp4
    %(prog)s input.mp4 output.mp4 --target-lufs -14
    %(prog)s input.mp4 output.mp4 --denoise
    %(prog)s input.mp4 output.mp4 --denoise --denoise-strength 0.7
    %(prog)s input.mp4 output.mp4 --verbose --denoise

  Batch directory mode:
    %(prog)s --input-dir ./videos --output-dir ./processed
    %(prog)s --input-dir ./videos --output-dir ./processed --denoise --target-lufs -14

How it works:
  1. Extracts audio from video to temporary WAV file
  2. Applies normalization and denoising to audio
  3. Replaces audio track in video (video stream copied, no re-encoding)
  4. Cleans up temporary files

Supported formats:
  MP4 with h264 or h265 video codec
  Video file must contain an audio track

LUFS (Loudness Units relative to Full Scale) is an industry standard for
measuring perceived loudness. Common target levels:
  -23 LUFS: Broadcast standard (EBU R128)
  -16 LUFS: General use, e-learning (recommended)
  -14 LUFS: Streaming platforms (Spotify, YouTube)

Denoising removes background noise (HVAC, room tone, ambient) before
normalization. Use --denoise-strength to control aggressiveness (0.0-1.0).
"""
    )

    # Create mutually exclusive group for single file vs batch mode
    input_group = parser.add_mutually_exclusive_group(required=True)

    input_group.add_argument(
        'input',
        nargs='?',
        type=Path,
        help='Input video file (single file mode)'
    )

    input_group.add_argument(
        '--input-dir',
        type=Path,
        help='Input directory containing video files (batch mode)'
    )

    parser.add_argument(
        'output',
        nargs='?',
        type=Path,
        help='Output video file (single file mode)'
    )

    parser.add_argument(
        '--output-dir',
        type=Path,
        help='Output directory for processed videos (batch mode)'
    )

    parser.add_argument(
        '--target-lufs',
        type=float,
        default=-16.0,
        metavar='FLOAT',
        help='Target loudness in LUFS (default: -16.0)'
    )

    parser.add_argument(
        '--denoise',
        action='store_true',
        help='Apply noise reduction before normalization (removes background noise like HVAC, room tone)'
    )

    parser.add_argument(
        '--denoise-strength',
        type=float,
        default=0.5,
        metavar='FLOAT',
        help='Denoising strength: 0.0 (no reduction) to 1.0 (maximum reduction) (default: 0.5)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )

    parser.add_argument(
        '-f', '--force',
        action='store_true',
        help='Overwrite output file without prompting'
    )

    args = parser.parse_args()

    # Validate mode: ensure either both input/output or both input-dir/output-dir are provided
    batch_mode = args.input_dir is not None
    single_mode = args.input is not None

    if batch_mode:
        if not args.output_dir:
            print("Error: --output-dir is required when using --input-dir")
            return 1
    else:
        if not args.output:
            print("Error: output file is required when using single file mode")
            return 1
        if args.output_dir:
            print("Error: --output-dir can only be used with --input-dir")
            return 1

    # Set up logging
    logger = setup_logging(args.verbose)

    # Banner
    logger.info("=" * 60)
    logger.info("Video Audio Normalizer")
    if batch_mode:
        logger.info("Batch Directory Mode")
    logger.info("=" * 60)

    # Validate target LUFS
    if args.target_lufs < -70 or args.target_lufs > 0:
        logger.error(f"Invalid target LUFS: {args.target_lufs}. Must be between -70 and 0.")
        return 1

    # Validate denoise strength
    if args.denoise_strength < 0.0 or args.denoise_strength > 1.0:
        logger.error(f"Invalid denoise strength: {args.denoise_strength}. Must be between 0.0 and 1.0.")
        return 1

    # Warn if denoise-strength specified without --denoise flag
    if args.denoise_strength != 0.5 and not args.denoise:
        logger.warning("--denoise-strength specified but --denoise flag not set. Denoising will not be applied.")

    # Check FFmpeg installation
    if not check_ffmpeg_installed():
        logger.error("FFmpeg is not installed or not accessible.")
        logger.error("FFmpeg is required for video audio extraction and replacement.")
        logger.error("")
        logger.error("Installation instructions:")
        logger.error("  Ubuntu/Debian: sudo apt-get install ffmpeg")
        logger.error("  macOS:         brew install ffmpeg")
        logger.error("  Windows:       Download from https://ffmpeg.org or use 'choco install ffmpeg'")
        return 1

    logger.info(f"FFmpeg: Available")
    logger.info(f"Target: {format_lufs(args.target_lufs)}")
    if args.denoise:
        logger.info(f"Denoise: Enabled (strength: {args.denoise_strength:.2f})")
    else:
        logger.info(f"Denoise: Disabled")
    logger.info("")

    try:
        if batch_mode:
            # Batch directory mode
            return process_batch_directory(args.input_dir, args.output_dir,
                                          args.target_lufs, args.force,
                                          args.denoise, args.denoise_strength, logger)
        else:
            # Single file mode
            logger.info(f"Input:  {args.input}")
            logger.info(f"Output: {args.output}")
            logger.info("")

            # Process video
            result = process_single_video(args.input, args.output, args.target_lufs,
                                        args.force, args.denoise, args.denoise_strength, logger)

            if result == 0:
                logger.info("")
                logger.info("=" * 60)
                logger.info("Processing Complete")
                logger.info("=" * 60)

            return result

    except KeyboardInterrupt:
        logger.info("")
        logger.warning("Operation cancelled by user.")
        return 130

    except Exception as e:
        logger.error("")
        logger.error(f"Fatal error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
