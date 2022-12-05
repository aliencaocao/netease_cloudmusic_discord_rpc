import gc
import wmi
from pyMeow import open_process, get_module, r_float64, close_process, get_module
import ctypes
import json
import sys
from ctypes import wintypes
from pypresence import Presence
import time
from win32com.client import Dispatch

__version__ = '0.1.2'
supported_cloudmusic_version = '2.10.6.3993'
current_offset = 0xA65568
maxlen_offset = 0xB16438  # TODO: does not work
print(f'网易云音乐Discord RPC v{__version__}，支持网易云音乐版本：{supported_cloudmusic_version}, modified by HackerRouter')

user32 = ctypes.windll.user32
WNDENUMPROC = ctypes.WINFUNCTYPE(
    wintypes.BOOL,
    wintypes.HWND,  # _In_ hWnd
    wintypes.LPARAM, )  # _In_ lParam

user32.EnumWindows.argtypes = (
    WNDENUMPROC,  # _In_ lpEnumFunc
    wintypes.LPARAM,)  # _In_ lParam

user32.IsWindowVisible.argtypes = (
    wintypes.HWND,)  # _In_ hWnd

user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowThreadProcessId.argtypes = (
    wintypes.HWND,  # _In_      hWnd
    wintypes.LPDWORD,)  # _Out_opt_ lpdwProcessId

user32.GetWindowTextLengthW.argtypes = (
    wintypes.HWND,)  # _In_ hWnd

user32.GetWindowTextW.argtypes = (
    wintypes.HWND,  # _In_  hWnd
    wintypes.LPWSTR,  # _Out_ lpString
    ctypes.c_int,)  # _In_  nMaxCount
wmic = wmi.WMI()

decoder = json.JSONDecoder()

def get_title(pid) -> str:
    ret = ''

    @WNDENUMPROC
    def enum_proc(hWnd, lParam):
        nonlocal ret
        if user32.IsWindowVisible(hWnd):
            _pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hWnd, ctypes.byref(_pid))
            length = user32.GetWindowTextLengthW(hWnd) + 1
            title = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hWnd, title, length)
            if _pid.value == pid: ret = title.value
        return True

    user32.EnumWindows(enum_proc, 0)
    return ret

def sec_to_str(sec) -> str:
    m, s = divmod(sec, 60)
    return f'{int(m):02d}:{int(s):02d}'

def get_playing(path):
    track_info = dict()
    with open(path, encoding='utf-8') as f:
        read_string = f.read(3200)
        for _ in range(4):
            try:
                read_string += f.read(500)
                decoded_json = decoder.raw_decode(read_string[1:])
                track_info.update(decoded_json[0])
                break
            except json.JSONDecodeError:
                pass
 
    if not track_info:
        return None

    picLink = track_info['track']['album']['picUrl']
    songID = track_info['track']['id']
    return songID, picLink

client_id = '1045242932128645180'
RPC = Presence(client_id)
RPC.connect()
print('RPC Launched\nThe following info will only be printed once for confirmation. They will continue to be updated to Discord.')
start_time = time.time()
first_run = True

while True:
    process = wmic.Win32_Process(name="cloudmusic.exe")
    process = [p for p in process if '--type=' not in p.ole_object.CommandLine]
    if not process:  # if the app isnt running, do nothing
        continue
    elif len(process) != 1:
        raise RuntimeError('Multiple candidate processes found!')
    else:
        process = process[0]
        ver_parser = Dispatch('Scripting.FileSystemObject')
        info = ver_parser.GetFileVersion(process.ExecutablePath)
        if info != supported_cloudmusic_version: raise RuntimeError(f'This version is not supported yet: {info}. Supported version: {supported_cloudmusic_version}')
        pid = process.ole_object.ProcessId
        if first_run: print(f'Found process: {pid}')
    song = get_title(pid)
    if not song: song = 'Unknown'
    process = open_process(pid)
    base_address = get_module(process, 'cloudmusic.dll')['base']
    current = r_float64(process, base_address + current_offset)
    current = sec_to_str(current)
    # maxlen = r_float64(process, base_address + maxlen_offset)  # not working now
    close_process(process)
    
    FilePath = "C:\\Users\\user\\AppData\\Local\\Netease\\CloudMusic\\webdata\\file\\history"
    songLinkPrefix = r"https://music.163.com/#/song?id="
    songLink, picUrl = get_playing(FilePath)
    songLink = songLinkPrefix + str(songLink)

    RPC.update(
        pid=pid, 
        details=f'Listening to {song}', 
        state=f'Currently at {current}', 
        large_image= picUrl, 
        large_text=song, 
        small_image= "logo",
        small_text="NetEase Cloud Music", 
        start=int(start_time),
        buttons = [{"label": "Play on browser", "url":songLink}, {"label": "Wanna know how it works?", "url":"https://github.com/HackerRouter/netease_cloudmusic_discord_rpc-modified"}]
        )

    if first_run: print(f'Song: {song}, current: {current}')
    first_run = False
    gc.collect()
    time.sleep(0.8)
