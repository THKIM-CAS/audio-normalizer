"""PPTX extraction and reconstruction functions."""

import logging
import os
import zipfile
from pathlib import Path
from typing import List, Tuple

logger = logging.getLogger('pptx_normalizer')


def extract_pptx(pptx_path: Path, temp_dir: Path) -> Tuple[Path, List[Path]]:
    """
    Extract PPTX file to temporary directory and find audio files.

    Args:
        pptx_path: Path to input PPTX file
        temp_dir: Temporary directory for extraction

    Returns:
        Tuple of (extraction_path, list of audio file paths)

    Raises:
        zipfile.BadZipFile: If PPTX is not a valid ZIP file
        Exception: For other extraction errors
    """
    logger.info(f"Extracting PPTX: {pptx_path}")

    # Extract PPTX contents
    extract_path = temp_dir / "pptx_contents"
    extract_path.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)

    logger.debug(f"Extracted to: {extract_path}")

    # Find audio files in ppt/media/ directory
    audio_files = find_audio_files(extract_path)

    logger.info(f"Found {len(audio_files)} audio file(s)")

    return extract_path, audio_files


def find_audio_files(extract_path: Path) -> List[Path]:
    """
    Find all audio files in the extracted PPTX.

    Args:
        extract_path: Path to extracted PPTX contents

    Returns:
        List of paths to audio files
    """
    audio_extensions = {'.wav', '.mp3', '.m4a', '.wma', '.aac', '.flac', '.ogg'}
    audio_files = []

    # Check ppt/media/ directory (standard location for PPTX media)
    media_dir = extract_path / "ppt" / "media"

    if media_dir.exists():
        for file_path in media_dir.iterdir():
            if file_path.is_file() and file_path.suffix.lower() in audio_extensions:
                audio_files.append(file_path)
                logger.debug(f"Found audio file: {file_path.name}")

    # Sort for consistent ordering
    audio_files.sort()

    return audio_files


def reconstruct_pptx(extract_path: Path, output_path: Path) -> None:
    """
    Reconstruct PPTX file from extracted and modified contents.

    Args:
        extract_path: Path to extracted PPTX contents
        output_path: Path for output PPTX file

    Raises:
        Exception: If reconstruction fails
    """
    logger.info(f"Reconstructing PPTX: {output_path}")

    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Create ZIP archive with DEFLATED compression
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        # Walk through all files and add them to the archive
        for root, dirs, files in os.walk(extract_path):
            for file in files:
                file_path = Path(root) / file

                # Calculate the archive name (relative path from extract_path)
                arcname = file_path.relative_to(extract_path)

                # Add file to archive
                zip_ref.write(file_path, arcname)
                logger.debug(f"Added to archive: {arcname}")

    logger.info(f"PPTX reconstruction complete: {output_path}")


def get_audio_format(file_path: Path) -> str:
    """
    Determine audio file format from extension.

    Args:
        file_path: Path to audio file

    Returns:
        Audio format as lowercase string (e.g., 'wav', 'mp3', 'm4a')
    """
    return file_path.suffix.lower().lstrip('.')
