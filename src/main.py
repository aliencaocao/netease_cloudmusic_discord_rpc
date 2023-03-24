import gc
import re
import sys
import time
import urllib.request
from os import path
from threading import Event, Thread
from typing import Callable

import orjson
import pythoncom
import wmi
from pyMeow import close_process, get_module, open_process, r_bytes, r_float64, r_uint
from pypresence import Presence
from win32com.client import Dispatch

__version__ = '0.2.1'
offsets = {'2.10.6.3993': {'current': 0xA65568, 'song_array': 0xB15654},
           '2.10.7.4239': {'current': 0xA66568, 'song_array': 0xB16974}}
interval = 1

if path.isfile('debug.log'):
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


def sec_to_str(sec: int) -> str:
    m, s = divmod(sec, 60)
    return f'{m:02d}:{s:02d}'


client_id = '1065646978672902144'
RPC = Presence(client_id)
RPC.connect()

print('RPC Launched.\nThe following info will only be printed only once for confirmation. They will continue being updated to Discord.')

start_time = time.time()
first_run = True
last_id = ''
last_int = 0

song_info_cache = {'': {'': ''}}


def get_song_info_from_netease(song_id: str) -> bool:
    try:
        song_info = {'': ''}
        url = f'https://music.163.com/song?id={song_id}'
        with urllib.request.urlopen(url) as response:
            html = response.read().decode('utf-8')

            re_img = re.compile(r'<meta property="og:image" content="(.+)"')
            song_info["cover"] = re_img.findall(html)[0]

            re_album = re.compile(r'<meta property="og:music:album" content="(.+)"')
            song_info["album"] = re_album.findall(html)[0]

            re_duration = re.compile(r'<meta property="music:duration" content="(.+)"')
            song_info["duration"] = re_duration.findall(html)[0]

            re_artist = re.compile(r'<meta property="og:music:artist" content="(.+)"')
            song_info["artist"] = re_artist.findall(html)[0]

            re_title = re.compile(r'<meta property="og:title" content="(.+)"')
            song_info["title"] = re_title.findall(html)[0]

            song_info_cache[song_id] = song_info
        return True
    except:
        print(f"Error while reading from remote: {song_id}")
        return False


def get_song_info_from_local(song_id: str) -> bool:
    filepath = path.join(path.expandvars('%LOCALAPPDATA%'), 'Netease/CloudMusic/webdata/file/history')
    if not path.exists(filepath):
        return False
    try:
        with(open(filepath, 'r', encoding='utf-8')) as f:
            history = orjson.loads(f.read())
            song_info_raw = next(x for x in history if str(x['track']['id']) == song_id)
            if not song_info_raw:
                return False
            song_info = {
                'cover': song_info_raw['track']['album']['picUrl'],
                'album': song_info_raw['track']['album']['name'],
                'duration': song_info_raw['track']['duration'] / 1000,
                'artist': '/'.join([x['name'] for x in song_info_raw['track']['artists']]),
                'title': song_info_raw['track']['name'],
            }
            song_info_cache[song_id] = song_info
        return True
    except:
        print(f"Error while reading from local history file: {song_id}")
        return False


def get_song_info(song_id: str) -> dict[str, str]:
    global song_info_cache
    if song_id not in song_info_cache:
        if not get_song_info_from_local(song_id):
            get_song_info_from_netease(song_id)
    return song_info_cache[song_id]


def update():
    global first_run
    global last_id
    global last_int

    try:
        pythoncom.CoInitialize()
        wmic = wmi.WMI()
        process = wmic.Win32_Process(name="cloudmusic.exe")
        process = [p for p in process if '--type=' not in p.ole_object.CommandLine]
        if not process:  # if the app isnt running, do nothing
            return
        elif len(process) != 1:
            raise RuntimeError('Multiple candidate processes found!')
        else:
            process = process[0]
            ver_parser = Dispatch('Scripting.FileSystemObject')
            ver = ver_parser.GetFileVersion(process.ExecutablePath)
            if ver not in offsets:
                raise RuntimeError(f'This version is not supported yet: {ver}. Supported version: {", ".join(offsets.keys())}')
            else:
                current_offset, song_array_offset = offsets[ver].values()
            pid = process.ole_object.ProcessId
            if first_run:
                print(f'Found process: {pid}')

        process = open_process(pid)
        module_base = get_module(process, 'cloudmusic.dll')['base']

        current_int = int(r_float64(process, module_base + current_offset))
        current_pystr = sec_to_str(current_int)

        songid_array = r_uint(process, module_base + song_array_offset)
        song_id = r_bytes(process, songid_array, 0x14).decode('utf-16')

        close_process(process)

        re_song_id = re.compile(r'(\d+)')
        if not re_song_id.match(song_id):
            # Song ID is not ready yet.
            return

        if song_id == last_id and current_int == last_int + interval:
            # Nothing changed.
            last_int += interval
            return

        song_info = get_song_info(song_id)

        try:
            RPC.update(pid=pid, details=f'{song_info["title"]}', state=f'{song_info["artist"]} | {song_info["album"]}', large_image=song_info["cover"],
                       large_text=song_info["album"].center(2), start=int(time.time() - current_int), 
                       buttons=[{"label": "Listen on Netease", "url": f"https://music.163.com/#/song?id={song_id}"}])
        except Exception as e:
            print("Error while updating Discord:", e)
            pass

        last_id = song_id
        last_int = current_int
        
        print(f'{song_info["title"]} - {song_info["artist"]}, {current_pystr}')

        first_run = False
        gc.collect()
        pythoncom.CoUninitialize()
    except Exception as e:
        print("Error while updating: ", e)


# calls the update function every second, ignore how long the actual update takes
RepeatedTimer(interval, update)
