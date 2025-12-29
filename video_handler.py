"""Video extraction and audio replacement functions."""

import json
import logging
import subprocess
from pathlib import Path
from typing import List, Optional, Tuple

logger = logging.getLogger('pptx_normalizer')


def extract_audio_from_video(video_path: Path, output_audio_path: Path) -> None:
    """
    Extract audio track from video file using FFmpeg.

    Args:
        video_path: Path to input video file
        output_audio_path: Path for extracted audio (WAV format)

    Raises:
        subprocess.CalledProcessError: If extraction fails
        ValueError: If video has no audio track
    """
    logger.info(f"Extracting audio from: {video_path.name}")

    # First check if video has audio stream
    if not _has_audio_stream(video_path):
        raise ValueError(f"Video file has no audio track: {video_path.name}")

    # Extract audio as 48kHz stereo WAV
    command = [
        'ffmpeg',
        '-i', str(video_path),
        '-vn',  # No video
        '-acodec', 'pcm_s16le',  # PCM 16-bit WAV
        '-ar', '48000',  # 48kHz sample rate
        '-ac', '2',  # Stereo
        '-y',  # Overwrite output file
        str(output_audio_path)
    ]

    try:
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True
        )
        logger.debug(f"Extracted audio to: {output_audio_path}")
    except subprocess.CalledProcessError as e:
        logger.error(f"FFmpeg audio extraction failed: {e.stderr}")
        raise


def replace_audio_in_video(video_path: Path, audio_path: Path,
                          output_video_path: Path) -> None:
    """
    Replace audio track in video file with processed audio.

    Args:
        video_path: Path to original video file
        audio_path: Path to processed audio (WAV format)
        output_video_path: Path for output video

    Raises:
        subprocess.CalledProcessError: If replacement fails
    """
    logger.info(f"Replacing audio in: {video_path.name}")

    # Replace audio: copy video stream, encode audio to AAC
    command = [
        'ffmpeg',
        '-i', str(video_path),
        '-i', str(audio_path),
        '-c:v', 'copy',  # Copy video stream (no re-encode)
        '-c:a', 'aac',  # Encode audio to AAC
        '-b:a', '192k',  # Audio bitrate
        '-map', '0:v:0',  # Use video from first input
        '-map', '1:a:0',  # Use audio from second input
        '-shortest',  # Match shortest stream duration
        '-y',  # Overwrite output file
        str(output_video_path)
    ]

    try:
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True
        )
        logger.debug(f"Created video with replaced audio: {output_video_path}")
    except subprocess.CalledProcessError as e:
        logger.error(f"FFmpeg audio replacement failed: {e.stderr}")
        raise


def validate_video_file(file_path: Path) -> Tuple[bool, Optional[str]]:
    """
    Validate video file exists and has supported format.

    Args:
        file_path: Path to video file

    Returns:
        Tuple of (is_valid, error_message)
    """
    # Check file exists
    if not file_path.exists():
        return False, f"File not found: {file_path}"

    # Check it's a file
    if not file_path.is_file():
        return False, f"Not a file: {file_path}"

    # Check extension
    if file_path.suffix.lower() != '.mp4':
        return False, f"Unsupported format: {file_path.suffix} (only .mp4 supported)"

    # Check for video and audio streams
    try:
        if not _has_video_stream(file_path):
            return False, f"No video stream found in: {file_path.name}"

        if not _has_audio_stream(file_path):
            return False, f"No audio stream found in: {file_path.name}"
    except Exception as e:
        return False, f"Failed to probe video file: {e}"

    return True, None


def find_video_files(directory: Path) -> List[Path]:
    """
    Find all MP4 video files in directory.

    Args:
        directory: Path to directory to search

    Returns:
        Sorted list of .mp4 file paths
    """
    if not directory.exists() or not directory.is_dir():
        return []

    video_files = []
    for file_path in directory.iterdir():
        if file_path.is_file() and file_path.suffix.lower() == '.mp4':
            video_files.append(file_path)
            logger.debug(f"Found video file: {file_path.name}")

    # Sort for consistent ordering
    video_files.sort()

    return video_files


def get_video_info(video_path: Path) -> dict:
    """
    Get video metadata using ffprobe.

    Args:
        video_path: Path to video file

    Returns:
        Dict with keys: duration, has_audio, has_video, video_codec, audio_codec

    Raises:
        subprocess.CalledProcessError: If ffprobe fails
    """
    command = [
        'ffprobe',
        '-v', 'error',
        '-show_entries', 'stream=codec_type,codec_name,duration',
        '-of', 'json',
        str(video_path)
    ]

    try:
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True
        )
        data = json.loads(result.stdout)

        info = {
            'duration': 0.0,
            'has_audio': False,
            'has_video': False,
            'video_codec': None,
            'audio_codec': None
        }

        for stream in data.get('streams', []):
            codec_type = stream.get('codec_type')
            codec_name = stream.get('codec_name')
            duration = stream.get('duration')

            if codec_type == 'video':
                info['has_video'] = True
                info['video_codec'] = codec_name
                if duration:
                    info['duration'] = max(info['duration'], float(duration))
            elif codec_type == 'audio':
                info['has_audio'] = True
                info['audio_codec'] = codec_name
                if duration:
                    info['duration'] = max(info['duration'], float(duration))

        return info

    except subprocess.CalledProcessError as e:
        logger.error(f"ffprobe failed: {e.stderr}")
        raise


def _has_audio_stream(video_path: Path) -> bool:
    """
    Check if video file has an audio stream.

    Args:
        video_path: Path to video file

    Returns:
        True if audio stream exists, False otherwise
    """
    try:
        info = get_video_info(video_path)
        return info['has_audio']
    except Exception:
        return False


def _has_video_stream(video_path: Path) -> bool:
    """
    Check if video file has a video stream.

    Args:
        video_path: Path to video file

    Returns:
        True if video stream exists, False otherwise
    """
    try:
        info = get_video_info(video_path)
        return info['has_video']
    except Exception:
        return False
