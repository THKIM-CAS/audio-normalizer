"""PPTX Sound Normalizer - Normalize audio volume across PowerPoint slides."""

__version__ = "0.1.0"

from audio_normalizer import normalize_audio_file, normalize_audio_files, AudioNormalizationStats
from pptx_handler import extract_pptx, reconstruct_pptx, find_audio_files, get_audio_format
from utils import (
    TempDirectory,
    setup_logging,
    check_ffmpeg_installed,
    validate_pptx_file,
    format_lufs,
    format_db,
)

__all__ = [
    "__version__",
    "normalize_audio_file",
    "normalize_audio_files",
    "AudioNormalizationStats",
    "extract_pptx",
    "reconstruct_pptx",
    "find_audio_files",
    "get_audio_format",
    "TempDirectory",
    "setup_logging",
    "check_ffmpeg_installed",
    "validate_pptx_file",
    "format_lufs",
    "format_db",
]
