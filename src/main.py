from __future__ import annotations

import asyncio
import gc
import os
import re
import sys
import time
from enum import IntFlag, auto
from threading import Event, Thread
from typing import Callable, TypedDict

import orjson
import pythoncom
import wmi
from pyMeow import close_process, get_module, get_process_name, open_process, pid_exists, r_bytes, r_float64, r_uint
from pyncm import apis
from pypresence import InvalidID, Presence
from win32api import GetFileVersionInfo, HIWORD, LOWORD

__version__ = '0.2.5'
offsets = {'2.7.1.1669': {'current': 0x8C8AF8, 'song_array': 0x8E9044},
           '2.10.5.3929': {'current': 0xA47548, 'song_array': 0xAF6FC8},
           '2.10.6.3993': {'current': 0xA65568, 'song_array': 0xB15654},
           '2.10.7.4239': {'current': 0xA66568, 'song_array': 0xB16974},
           '2.10.8.4337': {'current': 0xA74570, 'song_array': 0xB24f28}}
interval = 1

# regexes
re_song_id = re.compile(r'(\d+)')

if os.path.isfile('debug.log'):
    sys.stdout = open('debug.log', 'a')

print(f'Netease Cloud Music Discord RPC v{__version__}, Supporting NCM version: {", ".join(offsets.keys())}')


class RepeatedTimer:
    def __init__(self, interval: int, function: Callable[[], None]):
        self.interval = interval
        self.function = function
        self.start = time.time()
        self.event = Event()
        self.thread = Thread(target=self._target)
        self.thread.start()

    def _target(self):
        while not self.event.wait(self._time):
            self.function()

    @property
    def _time(self):
        return self.interval - ((time.time() - self.start) % self.interval)

    def stop(self):
        self.event.set()
        self.thread.join()


class SongInfo(TypedDict):
    cover: str
    album: str
    duration: float
    artist: str
    title: str


class Status(IntFlag):
    playing = auto()  # Song id unchanged and time += interval
    paused = auto()  # Song id unchanged and time unchanged
    changed = auto()  # Song id changed or time changed manually


def sec_to_str(sec: float) -> str:
    m, s = divmod(sec, 60)
    return f'{m:02.00f}:{s:05.02f}'


client_id = '1045242932128645180'
RPC = Presence(client_id)
RPC.connect()

print('RPC Launched.')

first_run = True
pid = 0
version = ''
last_status = Status.changed
last_id = ''
last_float = 0.0

song_info_cache: dict[str, SongInfo] = {}


def get_song_info_from_netease(song_id: str) -> bool:
    try:
        song_info_raw_list = apis.track.GetTrackDetail([song_id])['songs']
        if not song_info_raw_list:
            return False
        song_info_raw = song_info_raw_list[0]
        song_info: SongInfo = {
            'cover': song_info_raw['al']['picUrl'],
            'album': song_info_raw['al']['name'],
            'duration': song_info_raw['dt'] / 1000,
            'artist': ' / '.join([x['name'] for x in song_info_raw['ar']]),
            'title': song_info_raw['name']
        }
        song_info_cache[song_id] = song_info
        return True
    except Exception as e:
        print("Error while reading from remote:", e)
        return False


def get_song_info_from_local(song_id: str) -> bool:
    filepath = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Netease/CloudMusic/webdata/file/history')
    if not os.path.exists(filepath):
        return False
    try:
        with(open(filepath, 'r', encoding='utf-8')) as f:
            history = orjson.loads(f.read())
            song_info_raw_list = [x for x in history if str(x['track']['id']) == song_id]
            if not song_info_raw_list:
                return False
            song_info_raw = song_info_raw_list[0]['track']
            song_info: SongInfo = {
                'cover': song_info_raw['album']['picUrl'],
                'album': song_info_raw['album']['name'],
                'duration': song_info_raw['duration'] / 1000,
                'artist': '/'.join([x['name'] for x in song_info_raw['artists']]),
                'title': song_info_raw['name'],
            }
            song_info_cache[song_id] = song_info
        return True
    except Exception as e:
        print("Error while reading from local history file:", e)
        return False


def get_song_info(song_id: str) -> SongInfo:
    global song_info_cache
    if song_id not in song_info_cache:
        if not get_song_info_from_local(song_id):
            get_song_info_from_netease(song_id)
    return song_info_cache[song_id]


def find_process() -> tuple[int, str]:
    print('Searching for process...')
    pythoncom.CoInitialize()
    wmic = wmi.WMI()
    process_list = wmic.Win32_Process(name="cloudmusic.exe")
    process_list = [p for p in process_list if '--type=' not in p.CommandLine]
    if not process_list:
        return 0, ''
    if len(process_list) > 1:
        raise RuntimeError('Multiple candidate processes found!')
    process = process_list[0]

    ver_info = GetFileVersionInfo(process.ExecutablePath, '\\')
    ver = f"{HIWORD(ver_info['FileVersionMS'])}.{LOWORD(ver_info['FileVersionMS'])}." \
          f"{HIWORD(ver_info['FileVersionLS'])}.{LOWORD(ver_info['FileVersionLS'])}"

    pythoncom.CoUninitialize()
    return process.ProcessId, ver


def update():
    global first_run
    global pid
    global version
    global last_status
    global last_id
    global last_float

    try:
        if not pid_exists(pid) or get_process_name(pid) != 'cloudmusic.exe':
            pid, version = find_process()
            if not pid:
                # If the app isn't running, do nothing
                return

        if version not in offsets:
            raise RuntimeError(f'This version is not supported yet: {version}. Supported version: {", ".join(offsets.keys())}')
        current_offset, song_array_offset = offsets[version].values()

        if first_run:
            print(f'Found process: {pid}')
            first_run = False

        process = open_process(pid)
        module_base = get_module(process, 'cloudmusic.dll')['base']

        current_float = r_float64(process, module_base + current_offset)
        current_pystr = sec_to_str(current_float)

        songid_array = r_uint(process, module_base + song_array_offset)
        song_id = r_bytes(process, songid_array, 0x14).decode('utf-16').split('_')[0]  # Song ID can be shorter than 10 digits.

        close_process(process)

        if not re_song_id.match(song_id):
            # Song ID is not ready yet.
            return

        # Interval should fall in (interval - 0.2, interval + 0.2)
        status = Status.playing if song_id == last_id and abs(current_float - last_float - interval) < 0.1 else \
            Status.paused if song_id == last_id and current_float == last_float else \
                Status.changed

        if status == Status.playing:
            if last_status != Status.paused:  # Nothing changed
                last_float = current_float
                last_status = Status.playing
                return
        elif status == Status.paused:
            if last_status == Status.paused:  # Nothing changed
                return
            else:
                print('Paused')

        song_info = get_song_info(song_id)

        try:
            RPC.update(pid=pid,
                       state=f'{song_info["artist"]} | {song_info["album"]}',
                       details=song_info["title"].center(2),
                       large_image=song_info["cover"],
                       large_text=song_info["album"].center(2),
                       small_image='play' if status != Status.paused else 'pause',
                       small_text="Playing" if status != Status.paused else "Paused",
                       start=int(time.time() - current_float) if status != Status.paused else None,
                       buttons=[{"label": "Listen on Netease", "url": f"https://music.163.com/#/song?id={song_id}"}])
        except InvalidID:
            asyncio.set_event_loop(asyncio.new_event_loop())
            RPC.connect()
        except Exception as e:
            print("Error while updating Discord:", e)
            pass

        last_id = song_id
        last_float = current_float
        last_status = status

        if status != Status.paused:
            print(f'{song_info["title"]} - {song_info["artist"]}, {current_pystr}')

        gc.collect()
    except Exception as e:
        print("Error while updating: ", e)


# calls the update function every second, ignore how long the actual update takes
RepeatedTimer(interval, update)
