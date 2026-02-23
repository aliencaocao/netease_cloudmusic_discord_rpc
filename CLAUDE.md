# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Windows-only application that reads NetEase Cloud Music (NCM) playback state directly from process memory and displays it as Discord Rich Presence. Single-file Python app (`src/main.py`) with a Tkinter GUI and system tray icon.

## Commands

```bash
# Install dependencies (Python 3.10-3.13)
pip install -r src/requirements.txt

# Run in development
python src/main.py

# Run minimized to system tray
python src/main.py --min

# Build exe (run from repo root)
pyinstaller --log-level DEBUG --distpath ./ --clean --noconfirm src/main.spec

# Enable debug logging: create an empty file named debug.log in working directory
```

CI uses GitHub Actions (`.github/workflows/pyinstaller-windows.yml`): Python 3.12, PyInstaller + UPX on Windows.

## Architecture

The entire application lives in `src/main.py` (~517 lines). There are no tests, no separate modules.

### Core Loop

`startup()` → connects to Discord via pypresence → starts a `RepeatedTimer` (1-second interval) that calls `update()`:

1. **Process discovery**: Uses psutil to find `cloudmusic.exe`, reads file version via `win32api.GetFileVersionInfo`
2. **Memory reading**: Opens process with pyMeow, reads from `cloudmusic.dll`:
   - **V2.x**: base + hardcoded version-specific offsets (`current` → `r_float64` playback time, `song_array` → `r_uint` → UTF-16 song ID)
   - **V3.x**: AOB (array-of-bytes) pattern scan via `aob_scan_module()` to find `schedule_ptr` and `audio_player_ptr` at runtime (offsets change every launch). Song ID read via SSO string logic, UTF-8 encoded.
3. **Status detection** (line 354): Compares current song ID and playback time with previous values:
   - Playing: same ID, time advanced ~1s
   - Paused: same ID, time unchanged
   - Changed: different ID or manual seek
4. **Song info lookup**: Local history cache → playingList file → remote `pyncm` API. Results cached in `song_info_cache` dict.
5. **Discord update**: Sets presence with title, artist, album art URL, play/pause icon, elapsed time, and a "Listen on NetEase" button link.

### Key Design Details

- **Version-specific offsets** (line 40-52): Dict mapping NCM V2.x version strings to memory offsets. V3.x uses dynamic AOB scanning (`scan_for_v3_offsets()`) — patterns sourced from [Kxnrl/NetEase-Cloud-Music-DiscordRPC](https://github.com/Kxnrl/NetEase-Cloud-Music-DiscordRPC/blob/d3b77c679379aff1294cc83a285ad4f695376ad6/Vanessa/Players/NetEase.cs#L24).
- **Pause timeout**: After 30 minutes paused, disconnects RPC. Reconnects on resume.
- **Locale detection** (line 58): UI text is bilingual (Chinese/English) based on Windows UI language.
- **Startup on boot**: Writes/removes a `.bat` file in the Windows Startup folder.
- **`RepeatedTimer`** (line 77): Custom timer that compensates for execution time drift.

### Adding Support for New NCM Versions

**V2.x**: Use Cheat Engine to find the `current` (float64 playback time) and `song_array` offsets relative to `cloudmusic.dll` base. Add entry to the `offsets` dict with key format `major.minor.patch.build`.

**V3.x**: No action needed — all V3 versions are supported automatically via AOB pattern scanning. If NetEase changes their binary significantly, the byte patterns (`V3_AUDIO_PLAYER_PATTERN`, `V3_AUDIO_SCHEDULE_PATTERN`) may need updating.

## Dependencies

Key non-obvious dependencies:
- **pyMeow**: Windows process memory reading (installed from GitHub release zip, not PyPI)
- **pypresence**: Discord IPC Rich Presence client
- **pyncm**: Unofficial NetEase Cloud Music API wrapper (fallback for song metadata)
- **orjson**: Fast JSON parsing for local NCM history cache
- **pystray**: System tray icon (runs in daemon thread to avoid blocking the timer)

## Project Conventions

- All source code is in a single file `src/main.py` — no module structure
- Global mutable state for RPC connection, process info, and song cache
- Chinese comments and UI strings alongside English equivalents
- Discord Client ID: `1045242932128645180`
