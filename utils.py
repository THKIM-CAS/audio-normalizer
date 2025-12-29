"""Utility functions for PPTX audio normalizer."""

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional


class TempDirectory:
    """Context manager for temporary directory creation and cleanup."""

    def __init__(self, prefix: str = "pptx_normalizer_"):
        self.prefix = prefix
        self.path: Optional[Path] = None

    def __enter__(self) -> Path:
        """Create temporary directory."""
        self.path = Path(tempfile.mkdtemp(prefix=self.prefix))
        return self.path

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up temporary directory."""
        if self.path and self.path.exists():
            shutil.rmtree(self.path)


def setup_logging(verbose: bool = False) -> logging.Logger:
    """
    Set up logging configuration.

    Args:
        verbose: Enable verbose (DEBUG level) logging

    Returns:
        Configured logger instance
    """
    level = logging.DEBUG if verbose else logging.INFO

    logging.basicConfig(
        level=level,
        format='%(levelname)s: %(message)s'
    )

    return logging.getLogger('pptx_normalizer')


def check_ffmpeg_installed() -> bool:
    """
    Check if FFmpeg is installed and accessible.

    Returns:
        True if FFmpeg is available, False otherwise
    """
    try:
        result = subprocess.run(
            ['ffmpeg', '-version'],
            capture_output=True,
            text=True,
            timeout=5
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def validate_pptx_file(file_path: Path) -> tuple[bool, Optional[str]]:
    """
    Validate that a file exists and is a valid PPTX (ZIP) file.

    Args:
        file_path: Path to the PPTX file

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not file_path.exists():
        return False, f"File not found: {file_path}"

    if not file_path.is_file():
        return False, f"Not a file: {file_path}"

    # Check if it's a valid ZIP file (PPTX is a ZIP archive)
    try:
        import zipfile
        if not zipfile.is_zipfile(file_path):
            return False, f"Not a valid PPTX file (not a ZIP archive): {file_path}"
    except Exception as e:
        return False, f"Error validating file: {e}"

    return True, None


def format_lufs(lufs: float) -> str:
    """
    Format LUFS value for display.

    Args:
        lufs: LUFS value

    Returns:
        Formatted string
    """
    return f"{lufs:.1f} LUFS"


def format_db(db: float) -> str:
    """
    Format decibel value for display.

    Args:
        db: Decibel value

    Returns:
        Formatted string
    """
    sign = "+" if db > 0 else ""
    return f"{sign}{db:.1f} dB"


def get_audio_duration(sample_rate: int, num_samples: int) -> float:
    """
    Calculate audio duration in seconds.

    Args:
        sample_rate: Sample rate in Hz
        num_samples: Number of samples

    Returns:
        Duration in seconds
    """
    return num_samples / sample_rate


def confirm_overwrite(file_path: Path) -> bool:
    """
    Ask user to confirm overwriting an existing file.

    Args:
        file_path: Path to the file

    Returns:
        True if user confirms, False otherwise
    """
    response = input(f"Output file '{file_path}' already exists. Overwrite? [y/N]: ")
    return response.lower() in ('y', 'yes')


def find_pptx_files(directory: Path) -> list[Path]:
    """
    Find all PPTX files in a directory.

    Args:
        directory: Directory to search

    Returns:
        List of PPTX file paths, sorted by name
    """
    if not directory.exists():
        return []

    if not directory.is_dir():
        return []

    pptx_files = []
    for file_path in directory.iterdir():
        if file_path.is_file() and file_path.suffix.lower() == '.pptx':
            pptx_files.append(file_path)

    # Sort for consistent ordering
    pptx_files.sort()

    return pptx_files
