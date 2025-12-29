"""Audio normalization functions using LUFS standard."""

import logging
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Tuple

import numpy as np
import pyloudnorm as pyln
import soundfile as sf

from utils import get_audio_duration, format_lufs, format_db

logger = logging.getLogger('pptx_normalizer')


class AudioNormalizationStats:
    """Statistics for a normalized audio file."""

    def __init__(self, filename: str, original_lufs: float, target_lufs: float,
                 gain_db: float, duration: float):
        self.filename = filename
        self.original_lufs = original_lufs
        self.target_lufs = target_lufs
        self.gain_db = gain_db
        self.duration = duration

    def __str__(self) -> str:
        return (f"{self.filename}: {format_lufs(self.original_lufs)} → "
                f"{format_lufs(self.target_lufs)} ({format_db(self.gain_db)})")


def normalize_audio_file(file_path: Path, target_lufs: float = -16.0,
                        denoise: bool = False,
                        denoise_strength: float = 0.5) -> Optional[AudioNormalizationStats]:
    """
    Normalize a single audio file to target LUFS with optional denoising.

    Args:
        file_path: Path to audio file
        target_lufs: Target loudness in LUFS (default: -16.0)
        denoise: Whether to apply denoising before normalization (default: False)
        denoise_strength: Denoising strength 0.0-1.0 (default: 0.5)

    Returns:
        AudioNormalizationStats if successful, None if skipped/failed

    Raises:
        Exception: If normalization fails critically
    """
    logger.info(f"Processing: {file_path.name}")

    file_format = file_path.suffix.lower().lstrip('.')

    # Handle different audio formats
    if file_format in ['wav', 'flac', 'ogg']:
        # Native support via soundfile
        return _normalize_native_format(file_path, target_lufs, denoise, denoise_strength)
    elif file_format in ['mp3', 'm4a', 'wma', 'aac']:
        # Requires FFmpeg conversion
        return _normalize_compressed_format(file_path, target_lufs, file_format,
                                           denoise, denoise_strength)
    else:
        logger.warning(f"Unsupported audio format '{file_format}' for {file_path.name}, skipping")
        return None


def _normalize_native_format(file_path: Path, target_lufs: float,
                             denoise: bool = False,
                             denoise_strength: float = 0.5) -> Optional[AudioNormalizationStats]:
    """
    Normalize audio file in native format (WAV, FLAC, OGG).

    Args:
        file_path: Path to audio file
        target_lufs: Target loudness in LUFS
        denoise: Whether to apply denoising before normalization
        denoise_strength: Denoising strength (0.0-1.0)

    Returns:
        AudioNormalizationStats if successful, None if skipped
    """
    try:
        # Load audio file
        data, rate = sf.read(file_path)
        duration = get_audio_duration(rate, len(data))

        logger.debug(f"Loaded {file_path.name}: {rate} Hz, {duration:.2f}s, "
                    f"{data.shape[1] if len(data.shape) > 1 else 1} channel(s)")

        # Check minimum duration for LUFS measurement
        if duration < 0.4:
            logger.warning(f"Audio too short ({duration:.2f}s) for LUFS measurement, skipping {file_path.name}")
            return None

        # Apply denoising if requested
        if denoise:
            logger.debug(f"Applying denoising (strength: {denoise_strength:.2f})")
            try:
                data = denoise_audio(data, rate, denoise_strength)
                logger.debug(f"Denoising complete")
            except Exception as e:
                logger.warning(f"Denoising failed for {file_path.name}: {e}, continuing without denoising")
                # Continue with original audio if denoising fails

        # Measure loudness
        meter = pyln.Meter(rate)
        loudness = meter.integrated_loudness(data)

        logger.debug(f"Measured loudness: {format_lufs(loudness)}")

        # Check if audio is silent or measurement failed
        if loudness == float('-inf') or np.isnan(loudness):
            logger.warning(f"Cannot measure loudness (silent or invalid audio) for {file_path.name}, skipping")
            return None

        # Calculate gain
        gain_db = target_lufs - loudness

        # Normalize audio
        normalized_audio = pyln.normalize.loudness(data, loudness, target_lufs)

        # Write normalized audio back to file
        sf.write(file_path, normalized_audio, rate)

        logger.info(f"Normalized {file_path.name}: {format_lufs(loudness)} → {format_lufs(target_lufs)} "
                   f"({format_db(gain_db)})")

        return AudioNormalizationStats(
            filename=file_path.name,
            original_lufs=loudness,
            target_lufs=target_lufs,
            gain_db=gain_db,
            duration=duration
        )

    except Exception as e:
        logger.error(f"Failed to normalize {file_path.name}: {e}")
        return None


def _normalize_compressed_format(file_path: Path, target_lufs: float,
                                 original_format: str,
                                 denoise: bool = False,
                                 denoise_strength: float = 0.5) -> Optional[AudioNormalizationStats]:
    """
    Normalize compressed audio file (MP3, M4A, WMA) using FFmpeg conversion.

    Args:
        file_path: Path to audio file
        target_lufs: Target loudness in LUFS
        original_format: Original file format
        denoise: Whether to apply denoising before normalization
        denoise_strength: Denoising strength (0.0-1.0)

    Returns:
        AudioNormalizationStats if successful, None if skipped
    """
    temp_wav = None

    try:
        # Create temporary WAV file for processing
        with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as temp_file:
            temp_wav = Path(temp_file.name)

        # Convert to WAV using FFmpeg
        logger.debug(f"Converting {file_path.name} to WAV for processing")
        _convert_to_wav(file_path, temp_wav)

        # Load WAV file
        data, rate = sf.read(temp_wav)
        duration = get_audio_duration(rate, len(data))

        logger.debug(f"Loaded {file_path.name}: {rate} Hz, {duration:.2f}s")

        # Check minimum duration
        if duration < 0.4:
            logger.warning(f"Audio too short ({duration:.2f}s) for LUFS measurement, skipping {file_path.name}")
            return None

        # Apply denoising if requested
        if denoise:
            logger.debug(f"Applying denoising (strength: {denoise_strength:.2f})")
            try:
                data = denoise_audio(data, rate, denoise_strength)
                logger.debug(f"Denoising complete")
            except Exception as e:
                logger.warning(f"Denoising failed for {file_path.name}: {e}, continuing without denoising")

        # Measure loudness
        meter = pyln.Meter(rate)
        loudness = meter.integrated_loudness(data)

        logger.debug(f"Measured loudness: {format_lufs(loudness)}")

        # Check if audio is silent or measurement failed
        if loudness == float('-inf') or np.isnan(loudness):
            logger.warning(f"Cannot measure loudness (silent or invalid audio) for {file_path.name}, skipping")
            return None

        # Calculate gain
        gain_db = target_lufs - loudness

        # Normalize audio
        normalized_audio = pyln.normalize.loudness(data, loudness, target_lufs)

        # Write normalized audio to temp WAV
        sf.write(temp_wav, normalized_audio, rate)

        # Convert back to original format
        logger.debug(f"Converting normalized audio back to {original_format}")
        _convert_from_wav(temp_wav, file_path, original_format)

        logger.info(f"Normalized {file_path.name}: {format_lufs(loudness)} → {format_lufs(target_lufs)} "
                   f"({format_db(gain_db)})")

        return AudioNormalizationStats(
            filename=file_path.name,
            original_lufs=loudness,
            target_lufs=target_lufs,
            gain_db=gain_db,
            duration=duration
        )

    except Exception as e:
        logger.error(f"Failed to normalize {file_path.name}: {e}")
        return None

    finally:
        # Clean up temporary WAV file
        if temp_wav and temp_wav.exists():
            temp_wav.unlink()


def _convert_to_wav(input_path: Path, output_path: Path) -> None:
    """
    Convert audio file to WAV using FFmpeg.

    Args:
        input_path: Input audio file
        output_path: Output WAV file

    Raises:
        subprocess.CalledProcessError: If conversion fails
    """
    cmd = [
        'ffmpeg',
        '-i', str(input_path),
        '-y',  # Overwrite output file
        '-acodec', 'pcm_s16le',  # PCM 16-bit
        '-ar', '48000',  # 48 kHz sample rate
        '-ac', '2',  # Stereo
        '-loglevel', 'error',  # Only show errors
        str(output_path)
    ]

    subprocess.run(cmd, check=True, capture_output=True)


def _convert_from_wav(input_wav: Path, output_path: Path, target_format: str) -> None:
    """
    Convert WAV file back to target format using FFmpeg.

    Args:
        input_wav: Input WAV file
        output_path: Output audio file
        target_format: Target format (mp3, m4a, wma, etc.)

    Raises:
        subprocess.CalledProcessError: If conversion fails
    """
    # Format-specific codec settings
    codec_map = {
        'mp3': ['-acodec', 'libmp3lame', '-b:a', '192k'],
        'm4a': ['-acodec', 'aac', '-b:a', '192k'],
        'wma': ['-acodec', 'wmav2', '-b:a', '192k'],
        'aac': ['-acodec', 'aac', '-b:a', '192k'],
    }

    codec_args = codec_map.get(target_format, ['-acodec', 'copy'])

    cmd = [
        'ffmpeg',
        '-i', str(input_wav),
        '-y',  # Overwrite output file
        *codec_args,
        '-loglevel', 'error',  # Only show errors
        str(output_path)
    ]

    subprocess.run(cmd, check=True, capture_output=True)


def denoise_audio(data: np.ndarray, sample_rate: int, strength: float = 0.5) -> np.ndarray:
    """
    Remove background noise from audio using spectral gating.

    Args:
        data: Audio data as numpy array (mono or stereo)
        sample_rate: Sample rate in Hz
        strength: Denoising strength (0.0 = no reduction, 1.0 = maximum reduction)

    Returns:
        Denoised audio as numpy array with same shape as input

    Raises:
        Exception: If denoising fails
    """
    try:
        import noisereduce as nr
    except ImportError:
        raise ImportError(
            "noisereduce library not installed. Install with: uv pip install noisereduce"
        )

    # Ensure float32 dtype for noisereduce
    original_dtype = data.dtype
    if data.dtype != np.float32:
        data_float = data.astype(np.float32)
    else:
        data_float = data

    # Handle stereo (multi-channel) audio by processing each channel separately
    if len(data.shape) > 1 and data.shape[1] > 1:
        # Stereo or multi-channel audio
        num_channels = data.shape[1]
        denoised_channels = []

        for channel in range(num_channels):
            channel_data = data_float[:, channel]
            denoised_channel = nr.reduce_noise(
                y=channel_data,
                sr=sample_rate,
                stationary=True,
                prop_decrease=strength
            )
            denoised_channels.append(denoised_channel)

        # Stack channels back together
        reduced_noise = np.column_stack(denoised_channels)
    else:
        # Mono audio
        reduced_noise = nr.reduce_noise(
            y=data_float,
            sr=sample_rate,
            stationary=True,
            prop_decrease=strength
        )

    # Convert back to original dtype if needed
    if original_dtype != np.float32:
        reduced_noise = reduced_noise.astype(original_dtype)

    return reduced_noise


def normalize_audio_files(audio_files: list[Path], target_lufs: float = -16.0,
                          denoise: bool = False,
                          denoise_strength: float = 0.5) -> list[AudioNormalizationStats]:
    """
    Normalize multiple audio files with optional denoising.

    Args:
        audio_files: List of audio file paths
        target_lufs: Target loudness in LUFS
        denoise: Whether to apply denoising before normalization
        denoise_strength: Denoising strength 0.0-1.0

    Returns:
        List of AudioNormalizationStats for successfully normalized files
    """
    stats = []

    for audio_file in audio_files:
        result = normalize_audio_file(audio_file, target_lufs, denoise, denoise_strength)
        if result:
            stats.append(result)

    return stats
