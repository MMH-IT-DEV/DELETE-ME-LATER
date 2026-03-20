"""
Claude Code Stop Hook - Speaks the assistant's response using OpenAI TTS.
Reads the hook JSON from stdin, extracts the last assistant message,
summarizes it to a short spoken version, and plays it aloud.
"""

import sys
import json
import os
import re
import tempfile
import subprocess

def strip_markdown(text):
    """Remove markdown formatting, code blocks, and clean up for speech."""
    # Remove code blocks entirely (not useful to hear)
    text = re.sub(r'```[\s\S]*?```', '', text)
    # Remove inline code
    text = re.sub(r'`[^`]+`', '', text)
    # Remove markdown headers
    text = re.sub(r'#{1,6}\s*', '', text)
    # Remove markdown bold/italic
    text = re.sub(r'\*{1,3}([^*]+)\*{1,3}', r'\1', text)
    # Remove markdown links, keep text
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    # Remove markdown images
    text = re.sub(r'!\[([^\]]*)\]\([^)]+\)', '', text)
    # Remove HTML tags
    text = re.sub(r'<[^>]+>', '', text)
    # Remove markdown tables
    text = re.sub(r'\|[^\n]+\|', '', text)
    text = re.sub(r'[-|]+\s*\n', '', text)
    # Remove bullet points
    text = re.sub(r'^\s*[-*+]\s+', '', text, flags=re.MULTILINE)
    # Remove numbered lists prefix
    text = re.sub(r'^\s*\d+\.\s+', '', text, flags=re.MULTILINE)
    # Collapse whitespace
    text = re.sub(r'\n{2,}', '. ', text)
    text = re.sub(r'\n', ' ', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

def truncate_for_speech(text, max_chars=500):
    """Get a spoken summary - first meaningful sentences up to limit."""
    if len(text) <= max_chars:
        return text
    # Cut at sentence boundary
    truncated = text[:max_chars]
    last_period = truncated.rfind('.')
    last_question = truncated.rfind('?')
    last_exclaim = truncated.rfind('!')
    cut_point = max(last_period, last_question, last_exclaim)
    if cut_point > max_chars // 3:
        return truncated[:cut_point + 1]
    return truncated + "..."

def speak_with_openai(text, api_key):
    """Send text to OpenAI TTS and play the audio."""
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        # Save to temp file
        temp_path = os.path.join(tempfile.gettempdir(), "claude_tts.mp3")

        with client.audio.speech.with_streaming_response.create(
            model="tts-1",
            voice="shimmer",
            input=text,
            response_format="mp3"
        ) as response:
            response.stream_to_file(temp_path)

        # Play using Windows Media Player via PowerShell
        ps_script = f'''
        Add-Type -AssemblyName presentationCore
        $player = New-Object System.Windows.Media.MediaPlayer
        $player.Open([Uri]"{temp_path}")
        Start-Sleep -Milliseconds 300
        $player.Play()
        Start-Sleep -Milliseconds 500
        while ($player.Position -lt $player.NaturalDuration.TimeSpan) {{
            Start-Sleep -Milliseconds 200
        }}
        Start-Sleep -Milliseconds 200
        $player.Close()
        '''
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True, timeout=300
        )

        # Cleanup
        try:
            os.unlink(temp_path)
        except:
            pass

    except Exception as e:
        sys.stderr.write(f"TTS Error: {e}\n")

def main():
    try:
        hook_input = json.load(sys.stdin)
    except (json.JSONDecodeError, EOFError):
        return

    message = hook_input.get("last_assistant_message", "")
    if not message or len(message.strip()) < 10:
        return

    # Check if the stop hook is already active (prevent loops)
    if hook_input.get("stop_hook_active"):
        return

    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        sys.stderr.write("No OPENAI_API_KEY found\n")
        return

    cleaned = strip_markdown(message)
    if len(cleaned.strip()) < 5:
        return

    spoken = truncate_for_speech(cleaned, max_chars=8000)
    speak_with_openai(spoken, api_key)

if __name__ == "__main__":
    main()
