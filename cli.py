#!/usr/bin/env python3
"""
PPTX Narration Sound Tuner - Tune audio volume and remove noise from PowerPoint narrations.

This tool extracts audio from PPTX files, optionally removes background noise,
normalizes them to a target LUFS (Loudness Units relative to Full Scale) level,
and creates a new PPTX with tuned audio.
"""

import argparse
import sys
from pathlib import Path

from audio_normalizer import normalize_audio_files
from pptx_handler import extract_pptx, reconstruct_pptx
from utils import (
    TempDirectory,
    setup_logging,
    check_ffmpeg_installed,
    validate_pptx_file,
    confirm_overwrite,
    format_lufs,
    find_pptx_files
)


def process_single_pptx(input_path: Path, output_path: Path, target_lufs: float,
                        force: bool, denoise: bool, denoise_strength: float,
                        logger) -> int:
    """
    Process a single PPTX file.

    Args:
        input_path: Path to input PPTX file
        output_path: Path to output PPTX file
        target_lufs: Target loudness in LUFS
        force: Overwrite output without prompting
        denoise: Whether to apply denoising
        denoise_strength: Denoising strength (0.0-1.0)
        logger: Logger instance

    Returns:
        0 on success, 1 on error
    """
    # Validate input file
    is_valid, error_msg = validate_pptx_file(input_path)
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
            # Extract PPTX
            extract_path, audio_files = extract_pptx(input_path, temp_dir)

            if not audio_files:
                logger.warning("No audio files found in PPTX.")
                logger.info("Creating a copy of the original PPTX...")
                import shutil
                shutil.copy2(input_path, output_path)
                logger.info(f"Output written to: {output_path}")
                return 0

            # Normalize audio files
            stats = normalize_audio_files(audio_files, target_lufs, denoise, denoise_strength)

            # Display summary
            if stats:
                logger.info(f"Normalized {len(stats)} of {len(audio_files)} audio file(s)")
            else:
                logger.warning("No audio files were normalized.")

            if len(stats) < len(audio_files):
                skipped = len(audio_files) - len(stats)
                logger.warning(f"Skipped {skipped} file(s) due to errors or unsupported formats.")

            # Reconstruct PPTX
            reconstruct_pptx(extract_path, output_path)
            logger.info(f"Success! Output written to: {output_path}")

            return 0

    except Exception as e:
        logger.error(f"Failed to process {input_path.name}: {e}")
        return 1


def main():
    """Main entry point for the PPTX narration sound tuner."""
    parser = argparse.ArgumentParser(
        description='Tune audio volume and remove noise from PowerPoint narrations using LUFS standard.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Single file mode:
    %(prog)s input.pptx output.pptx
    %(prog)s input.pptx output.pptx --target-lufs -14
    %(prog)s input.pptx output.pptx --denoise
    %(prog)s input.pptx output.pptx --denoise --denoise-strength 0.7
    %(prog)s input.pptx output.pptx --verbose --denoise

  Batch directory mode:
    %(prog)s --input-dir ./input --output-dir ./output
    %(prog)s --input-dir ./input --output-dir ./output --denoise --target-lufs -14

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
        help='Input PPTX file (single file mode)'
    )

    input_group.add_argument(
        '--input-dir',
        type=Path,
        help='Input directory containing PPTX files (batch mode)'
    )

    parser.add_argument(
        'output',
        nargs='?',
        type=Path,
        help='Output PPTX file (single file mode)'
    )

    parser.add_argument(
        '--output-dir',
        type=Path,
        help='Output directory for normalized PPTX files (batch mode)'
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
    logger.info("PPTX Narration Sound Tuner")
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
        logger.error("FFmpeg is required for processing compressed audio formats (MP3, M4A, etc.).")
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

            # Validate input file
            is_valid, error_msg = validate_pptx_file(args.input)
            if not is_valid:
                logger.error(error_msg)
                return 1

            # Check if output file exists
            if args.output.exists() and not args.force:
                if not confirm_overwrite(args.output):
                    logger.info("Operation cancelled.")
                    return 0

            # Process PPTX
            with TempDirectory() as temp_dir:
                # Extract PPTX
                extract_path, audio_files = extract_pptx(args.input, temp_dir)

                if not audio_files:
                    logger.warning("No audio files found in PPTX.")
                    logger.info("Creating a copy of the original PPTX...")
                    import shutil
                    shutil.copy2(args.input, args.output)
                    logger.info(f"Output written to: {args.output}")
                    return 0

                logger.info("")

                # Normalize audio files
                stats = normalize_audio_files(audio_files, args.target_lufs,
                                              args.denoise, args.denoise_strength)

                logger.info("")
                logger.info("-" * 60)

                # Display summary
                if stats:
                    logger.info(f"Successfully normalized {len(stats)} of {len(audio_files)} audio file(s):")
                    logger.info("")
                    for stat in stats:
                        logger.info(f"  {stat}")
                else:
                    logger.warning("No audio files were normalized.")

                if len(stats) < len(audio_files):
                    skipped = len(audio_files) - len(stats)
                    logger.warning(f"Skipped {skipped} file(s) due to errors or unsupported formats.")

                logger.info("")
                logger.info("-" * 60)

                # Reconstruct PPTX
                reconstruct_pptx(extract_path, args.output)

                logger.info("")
                logger.info("=" * 60)
                logger.info(f"Success! Output written to: {args.output}")
                logger.info("=" * 60)

                return 0

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


def process_batch_directory(input_dir: Path, output_dir: Path, target_lufs: float,
                            force: bool, denoise: bool, denoise_strength: float,
                            logger) -> int:
    """
    Process all PPTX files in a directory.

    Args:
        input_dir: Input directory containing PPTX files
        output_dir: Output directory for normalized PPTX files
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

    # Find all PPTX files
    pptx_files = find_pptx_files(input_dir)

    if not pptx_files:
        logger.warning(f"No PPTX files found in: {input_dir}")
        return 0

    logger.info(f"Input Directory:  {input_dir}")
    logger.info(f"Output Directory: {output_dir}")
    logger.info(f"Found {len(pptx_files)} PPTX file(s) to process")
    logger.info("")

    # Create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Process each PPTX file
    success_count = 0
    error_count = 0

    for i, input_file in enumerate(pptx_files, 1):
        logger.info("=" * 60)
        logger.info(f"Processing file {i}/{len(pptx_files)}: {input_file.name}")
        logger.info("=" * 60)

        output_file = output_dir / input_file.name

        result = process_single_pptx(input_file, output_file, target_lufs, force,
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
    logger.info(f"Total files:      {len(pptx_files)}")
    logger.info(f"Successfully processed: {success_count}")
    if error_count > 0:
        logger.info(f"Errors:           {error_count}")
    logger.info("")

    return 0 if error_count == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
