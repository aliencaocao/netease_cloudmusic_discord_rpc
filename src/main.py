import ctypes
import locale
import logging
import os
import re
import sys
import time
import webbrowser
from enum import IntFlag, auto
from threading import Event as ThreadingEvent, Thread
from tkinter import *
from tkinter import BooleanVar, messagebox
from tkinter.ttk import *
from typing import Callable, Dict, Tuple, TypedDict

import orjson
import psutil
from PIL import Image
from pyMeow import aob_scan_module, close_process, get_module, get_process_name, open_process, pid_exists, r_bytes, r_float64, r_int, r_int64, r_uint
from pyncm import apis
from pypresence import DiscordNotFound, PipeClosed, Presence
from pystray import Icon as TrayIcon, Menu as TrayMenu, MenuItem as TrayItem
from win32api import GetFileVersionInfo, HIWORD, LOWORD

__version__ = '0.3.6'

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)
if os.path.isfile('debug.log'):
    file_handler = logging.FileHandler('debug.log', encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

offsets = {
    '2.7.1.1669': {'current': 0x8C8AF8, 'song_array': 0x8E9044},
    '2.10.3.3613': {'current': 0xA39550, 'song_array': 0xAE8F80},
    '2.10.5.3929': {'current': 0xA47548, 'song_array': 0xAF6FC8},
    '2.10.6.3993': {'current': 0xA65568, 'song_array': 0xB15654},
    '2.10.7.4239': {'current': 0xA66568, 'song_array': 0xB16974},
    '2.10.8.4337': {'current': 0xA74570, 'song_array': 0xB24F28},
    '2.10.10.4509': {'current': 0xA77580, 'song_array': 0xB282CC},
    '2.10.10.4689': {'current': 0xA79580, 'song_array': 0xB2AD10},
    '2.10.11.4930': {'current': 0xA7A580, 'song_array': 0xB2BCB0},
    '2.10.12.5241': {'current': 0xA7A580, 'song_array': 0xB2BCB0},
    '2.10.13.6067': {'current': 0xA7A590, 'song_array': 0xB2BCD0},
}
# V3 byte patterns for dynamic memory scanning (offsets change every launch)
# Source: https://github.com/Kxnrl/NetEase-Cloud-Music-DiscordRPC/blob/d3b77c679379aff1294cc83a285ad4f695376ad6/Vanessa/Players/NetEase.cs#L24
V3_AUDIO_PLAYER_PATTERN = "48 8D 0D ?? ?? ?? ?? E8 ?? ?? ?? ?? 48 8D 0D ?? ?? ?? ?? E8 ?? ?? ?? ?? 90 48 8D 0D ?? ?? ?? ?? E8 ?? ?? ?? ?? 48 8D 05 ?? ?? ?? ?? 48 8D A5 ?? ?? ?? ?? 5F 5D C3 CC CC CC CC CC 48 89 4C 24 ?? 55 57 48 81 EC ?? ?? ?? ?? 48 8D 6C 24 ?? 48 8D 7C 24"
V3_AUDIO_SCHEDULE_PATTERN = "66 0F 2E 0D ?? ?? ?? ?? 7A ?? 75 ?? 66 0F 2E 15"

# Cached V3 pointers (resolved per process launch via AOB scan)
v3_schedule_ptr = 0
v3_audio_player_ptr = 0

frozen = getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')
interval = 1
windll = ctypes.windll.kernel32
is_CN = locale.windows_locale[windll.GetUserDefaultUILanguage()].startswith('zh_')
user_startup_folder = os.path.join(os.path.expandvars('%APPDATA%'), r'Microsoft\Windows\Start Menu\Programs\Startup')
startup_file_path = os.path.join(user_startup_folder, 'Netease Cloud Music Discord RPC.bat')
start_minimized = '--min' in sys.argv
re_song_id = re.compile(r'(\d+)')

logger.info(f"Netease Cloud Music Discord RPC v{__version__}\nRunning on Python {sys.version}\nSupporting NCM version: {', '.join(offsets.keys())}, 3.x (dynamic scan)")


def get_res_path(relative_path: str) -> str:
    """ Get absolute path to resource, works for dev and for PyInstaller
     Relative path will always get extracted into root!"""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    if os.path.exists(os.path.join(base_path, relative_path)):
        return os.path.join(base_path, relative_path)
    else:
        raise FileNotFoundError(f'{os.path.join(base_path, relative_path)} is not found!')


class RepeatedTimer:
    def __init__(self, interval: int, function: Callable[[], None], stop_variable: ThreadingEvent):
        self.interval = interval
        self.function = function
        self.stop_variable = stop_variable
        self.start = time.time()
        self.event = ThreadingEvent()
        self.thread = Thread(target=self._target)
        self.thread.start()

    def _target(self):
        while not self.stop_variable.is_set() and not self.event.wait(self._time):
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


class UnsupportedVersionError(Exception):
    pass


def sec_to_str(sec: float) -> str:
    m, s = divmod(sec, 60)
    return f'{m:02.00f}:{s:05.02f}'


def connect_discord(presence: Presence) -> bool:
    global connected
    try:
        presence.connect()
    except DiscordNotFound:
        connected = False
        logger.warning('Discord not found.')
        if not start_minimized:
            root.after(0, lambda: messagebox.showerror('Discord not found', 'Could not detect a running Discord instance. Please make sure Discord is running and try again. Do not use BetterDiscord or other 3rd party clients.'))
        return False
    except Exception as e:
        connected = False
        logger.warning('Error while connecting to Discord:', e)
        return False
    else:
        connected = True
        logger.info('Discord Connected.')
        return True


def disconnect_discord(presence: Presence):
    global connected
    try:
        presence.clear()
        presence.close()
    except Exception as e:
        logger.warning(f'Error while disconnecting Discord:', e)
        connected = False  # set to false anyways because the only reason why it could fail is due to already disconnected/already closed async loop, which means it is disconnected already
        return False
    else:
        logger.info(f'Disconnected from Discord.')
        connected = False
        return True


client_id = '1045242932128645180'
RPC = Presence(client_id)

first_run = True
pid = 0
version = ''
last_status = Status.changed
last_id = ''
last_float = 0.0
last_pause_time = time.time()
pause_timeout = 30 * 60
stop_variable = ThreadingEvent()

song_info_cache: Dict[str, SongInfo] = {}
cached_process = None  # pyMeow process handle, reused across ticks
cached_module_base = 0  # V2 cloudmusic.dll base address, stable per process
connected = False  # Discord RPC connection state (plain bool, not BooleanVar — thread-safe under GIL)


def toggle():
    global timer
    menu = icon.menu
    if not toggle_var.get():
        if connect_discord(RPC):
            stop_variable.clear()
            timer = RepeatedTimer(interval, update, stop_variable=stop_variable)
            toggle_var.set(True)
            menu = TrayMenu(*[disable_item] + org_menu)
    else:
        stop_update()
        toggle_var.set(False)
        menu = TrayMenu(*[enable_item] + org_menu)
    icon.menu = menu


def toggle_startup():
    if startup_var.get():
        logger.debug('Adding startup file')
        with open(startup_file_path, 'w+') as f:
            if frozen:
                cmd = f'start "" "{sys.executable}" --min'
            else:
                cmd = f'start "" "{sys.executable}" "{__file__}" --min'
            logger.debug(f'Writing to {os.path.join(user_startup_folder, "Netease Cloud Music Discord RPC.bat")}\n{cmd}')
            f.write(cmd)
    else:
        logger.debug('Removing startup file')
        if os.path.isfile(startup_file_path):
            os.remove(startup_file_path)


def about():
    supported_ver_str = '\n'.join(offsets.keys()) + '\n3.x (dynamic scan)'
    messagebox.showinfo('About', f"Netease Cloud Music Discord RPC v{__version__}\nPython {sys.version}\nSupporting NCM version:\n{supported_ver_str}\nMaintainer: Billy Cao" if not is_CN else
    f"网易云音乐 Discord RPC v{__version__}\nPython版本 {sys.version}\n支持的网易云音乐版本:\n{supported_ver_str}\n开发者: Billy Cao")


def quit_app(icon=None, item=None):
    stop_update()
    if icon: icon.stop()
    root.destroy()


def show_window(icon, item):
    icon.stop()
    root.after(0, root.deiconify)


def hide_window():
    global icon
    root.withdraw()
    if toggle_var.get():
        menu = TrayMenu(*[disable_item] + org_menu)
    else:
        menu = TrayMenu(*[enable_item] + org_menu)
    icon = TrayIcon("Netease Cloud Music Discord RPC", icon_image, "Netease Cloud Music Discord RPC", menu)
    Thread(target=icon.run, daemon=True).start()  # must run this in thread, else it block the update RepeatTimer thread


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
            'artist': '/'.join([x['name'] for x in song_info_raw['ar']]),
            'title': song_info_raw['name'],
        }
        song_info_cache[song_id] = song_info
        return True
    except Exception as e:  # normal to fail when playing a cloud drive uploaded file since song ID is not public
        logger.warning('Error while reading from remote:', e)
        return False


def get_song_info_from_local(song_id: str) -> bool:
    filepath = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Netease/CloudMusic/webdata/file/history')
    if not os.path.exists(filepath):
        return False
    try:
        with (open(filepath, 'r', encoding='utf-8')) as f:
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
        logger.warning('Error while reading from local history file:', e)
        return False


def get_song_info_from_playing_list(song_id: str) -> bool:
    filepath = os.path.join(os.path.expandvars('%LOCALAPPDATA%'), 'Netease/CloudMusic/WebData/file/playingList')
    if not os.path.exists(filepath):
        return False
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = orjson.loads(f.read())
            song_list = data.get('list', [])
            song_info_raw_list = [x for x in song_list if str(x.get('id', '')) == song_id]
            if not song_info_raw_list:
                return False
            track = song_info_raw_list[0]['track']
            song_info: SongInfo = {
                'cover': track['album']['cover'],
                'album': track['album']['name'],
                'duration': track.get('duration', 0) / 1000 if track.get('duration', 0) else 0,
                'artist': '/'.join([x['name'] for x in track['artists']]),
                'title': track['name'],
            }
            song_info_cache[song_id] = song_info
        return True
    except Exception as e:
        logger.warning('Error while reading from playingList file:', e)
        return False


def get_song_info(song_id: str) -> SongInfo | None:
    if song_id not in song_info_cache:
        if not get_song_info_from_local(song_id):
            if not get_song_info_from_playing_list(song_id):
                get_song_info_from_netease(song_id)
    return song_info_cache.get(song_id)


def find_process() -> Tuple[int, str]:
    candidates = []
    for proc in psutil.process_iter(attrs=['name', 'pid']):
        if proc.info['name'] == 'cloudmusic.exe':
            try:
                cmdline = proc.cmdline()
                if any('--type=' in arg for arg in cmdline):
                    continue
                candidates.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    if not candidates:
        return 0, ''
    if len(candidates) > 1:
        raise RuntimeError('Multiple candidate processes found!')
    proc = candidates[0]
    try:
        exe_path = proc.exe()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return 0, ''

    ver_info = GetFileVersionInfo(exe_path, '\\')
    ver = (f"{HIWORD(ver_info['FileVersionMS'])}.{LOWORD(ver_info['FileVersionMS'])}."
           f"{HIWORD(ver_info['FileVersionLS'])}.{LOWORD(ver_info['FileVersionLS'])}")
    return proc.info['pid'], ver


def scan_for_v3_offsets(process: dict, module_name: str = 'cloudmusic.dll') -> Tuple[int, int]:
    """Scan cloudmusic.dll for V3 audio pointers using AOB patterns.
    Returns (schedule_ptr, audio_player_ptr) as absolute virtual addresses."""
    results = aob_scan_module(process, module_name, V3_AUDIO_SCHEDULE_PATTERN)
    if not results:
        raise RuntimeError('V3 AOB scan failed: AudioSchedulePattern not found')
    match = results[0]
    text_addr = match + 4
    displacement = r_int(process, text_addr)
    schedule_ptr = text_addr + displacement + 4
    logger.debug(f'V3 schedule pointer: {hex(schedule_ptr)}')

    results = aob_scan_module(process, module_name, V3_AUDIO_PLAYER_PATTERN)
    if not results:
        raise RuntimeError('V3 AOB scan failed: AudioPlayerPattern not found')
    match = results[0]
    text_addr = match + 3
    displacement = r_int(process, text_addr)
    audio_player_ptr = text_addr + displacement + 4
    logger.debug(f'V3 audio player pointer: {hex(audio_player_ptr)}')

    return schedule_ptr, audio_player_ptr


def read_v3_song_id(process: dict, audio_player_ptr: int) -> str:
    """Read current song ID from V3 memory layout (UTF-8, SSO string)."""
    audio_play_info = r_int64(process, audio_player_ptr + 0x50)
    if audio_play_info == 0:
        return ''

    str_ptr = audio_play_info + 0x10
    str_length = r_int64(process, str_ptr + 0x10)

    if str_length <= 0:
        return ''

    # Cap read size to avoid reading excessive memory; song ID strings are short (e.g. "1234567890_0")
    read_length = min(int(str_length), 128)

    # Small string optimization: if length <= 15, data is inline at str_ptr; otherwise dereference
    if str_length <= 15:
        raw = r_bytes(process, str_ptr, read_length)
    else:
        str_address = r_int64(process, str_ptr)
        if str_address == 0:
            return ''
        raw = r_bytes(process, str_address, read_length)

    song_str = raw.decode('utf-8')
    if not song_str or '_' not in song_str:
        return ''
    return song_str[:song_str.index('_')]


def update():
    global first_run
    global pid
    global version
    global last_status
    global last_id
    global last_float
    global last_pause_time
    global v3_schedule_ptr
    global v3_audio_player_ptr
    global cached_process
    global cached_module_base
    global connected

    try:
        if not pid_exists(pid) or get_process_name(pid) != 'cloudmusic.exe':
            # Process died or changed — close cached handle and reset
            if cached_process is not None:
                close_process(cached_process)
                cached_process = None
                cached_module_base = 0
            pid, version = find_process()
            if not pid:  # If netease client isn't running, clear presence
                logger.warning('Netease Cloud Music not found.')
                root.after(0, lambda: song_title_text.set('N/A'))
                root.after(0, lambda: song_artist_text.set(''))
                if not first_run:
                    disconnect_discord(RPC)
                    first_run = True
                return
            first_run = True  # New PID found — trigger re-scan of offsets

        is_v3 = version.startswith('3.')
        if not is_v3 and version not in offsets:
            stop_variable.set()
            raise UnsupportedVersionError(f"This version is not supported yet: {version}.\nSupported version: {', '.join(offsets.keys())}" if not is_CN else f"目前不支持此网易云音乐版本: {version}。\n支持的版本: {', '.join(offsets.keys())}")

        # Reuse cached process handle; open only when needed
        if cached_process is None:
            cached_process = open_process(pid)

        if first_run:
            logger.info(f'Found process: {pid}')
            if is_v3:
                v3_schedule_ptr, v3_audio_player_ptr = scan_for_v3_offsets(cached_process, 'cloudmusic.dll')
                logger.info(f'V3 AOB scan complete: schedule={hex(v3_schedule_ptr)}, player={hex(v3_audio_player_ptr)}')
            else:
                cached_module_base = get_module(cached_process, 'cloudmusic.dll')['base']
            first_run = False

        if is_v3:
            current_float = r_float64(cached_process, v3_schedule_ptr)
            song_id = read_v3_song_id(cached_process, v3_audio_player_ptr)
        else:
            current_float = r_float64(cached_process, cached_module_base + offsets[version]['current'])
            songid_array = r_uint(cached_process, cached_module_base + offsets[version]['song_array'])
            song_id = r_bytes(cached_process, songid_array, 0x14).decode('utf-16').split('_')[0]  # Song ID can be shorter than 10 digits.

        current_pystr = sec_to_str(current_float)

        if not re_song_id.match(song_id):
            # Song ID is not ready yet.
            return

        # Measured interval should fall in (interval +- 0.2)
        status = (Status.playing if song_id == last_id and abs(current_float - last_float - interval) < 0.2
                  else Status.paused if song_id == last_id and current_float == last_float
        else Status.changed)
        if status == Status.playing:
            if last_status != Status.paused:  # Nothing changed
                last_float = current_float
                last_status = Status.playing
                return
            elif last_status == Status.paused:  # we resumed from pause and may need to reconnect if passed time out
                logger.debug('Resumed')
                if not connected:
                    if not connect_discord(RPC):
                        return

        elif status == Status.paused:
            if last_status == Status.paused:  # Nothing changed but check if it is idle/paused for more than 30min, clear presence but keep connection alive to avoid reconnection throttling.
                if connected and time.time() - last_pause_time > pause_timeout:
                    logger.debug('Idle for more than 30min, clearing RPC presence.')
                    try:
                        RPC.clear()
                    except Exception:
                        pass
                return
            elif last_status == Status.playing:
                logger.debug('Paused')
                last_pause_time = time.time()

        elif status == Status.changed:
            last_pause_time = time.time()  # reset timeout as changed indicates something happened
            if not connected:
                if not connect_discord(RPC):
                    return

        song_info = get_song_info(song_id)

        if song_info is None:
            logger.warning(f'Could not find song info for ID: {song_id}')
            # Still advance tracking state to avoid infinite Status.changed loop
            last_id = song_id
            last_float = current_float
            return

        if not (connected or not time.time() - last_pause_time > pause_timeout):
            if not connect_discord(RPC):
                return

        try:
            RPC.update(pid=pid,
                       state=f"{song_info['artist']} | {song_info['album']}",
                       details=song_info['title'].center(2),
                       large_image=song_info['cover'],
                       large_text=song_info['album'].center(2),
                       small_image='play' if status != Status.paused else 'pause',
                       small_text='Playing' if status != Status.paused else 'Paused',
                       start=int(time.time() - current_float)
                       if status != Status.paused else None,
                       # Known issue: buttons do not appear on PC due to discord API changes: https://github.com/qwertyquerty/pypresence/issues/237
                       buttons=[{'label': 'Listen on NetEase',
                                 'url': f'https://music.163.com/#/song?id={song_id}'}]
                       )
            title = song_info['title']
            artist = song_info['artist']
            root.after(0, lambda t=title, a=artist: (song_title_text.set(t), song_artist_text.set(a)))
        except PipeClosed:
            logger.info('Reconnecting to Discord...')
            connect_discord(RPC)
        except Exception as e:
            logger.error('Error while updating to Discord:')
            logger.exception(e)

        last_id = song_id
        last_float = current_float
        if status != Status.changed:  # only store play/pause status for ease of detection above
            last_status = status

        if status != Status.paused:
            logger.debug(f"{song_info['title']} - {song_info['artist']}, {current_pystr}")
    except UnsupportedVersionError as e:
        if not start_minimized:
            msg = str(e)
            root.after(0, lambda m=msg: messagebox.showerror('不支持的网易云音乐版本', m))
        root.after(0, lambda: toggle_var.set(False))
        stop_variable.set()
        raise e
    except Exception as e:
        logger.error('Error while updating song info:')
        logger.exception(e)


def startup():
    global timer
    if start_minimized:
        hide_window()
    if connect_discord(RPC):
        try:
            timer = RepeatedTimer(interval, update, stop_variable=stop_variable)
        except UnsupportedVersionError:
            return  # handled in update() already
        except Exception as e:
            messagebox.showerror('Error' if not is_CN else '错误', f'{e}')
            toggle_var.set(False)
            stop_variable.set()
            return
        else:
            toggle_var.set(True)
    else:
        toggle_var.set(False)


def stop_update():
    global cached_process, cached_module_base
    stop_variable.set()
    if 'timer' in globals():
        timer.stop()
    if cached_process is not None:
        close_process(cached_process)
        cached_process = None
        cached_module_base = 0
    try:
        RPC.clear()
        RPC.close()
    except:  # if not connected then it will error, just ignore it
        pass


org_menu = [TrayItem('Show' if not is_CN else '显示主窗口', show_window, default=True),
            TrayItem('Quit' if not is_CN else '退出', quit_app)]
enable_item = TrayItem('Enable' if not is_CN else '启用', toggle)
disable_item = TrayItem('Disable' if not is_CN else '禁用', toggle)
icon_image = Image.open(get_res_path("app_logo.png"))
icon = TrayIcon("Netease Cloud Music Discord RPC", icon_image, "Netease Cloud Music Discord RPC", org_menu)

root = Tk()
root.title('Netease Cloud Music Discord RPC')
root.resizable(False, False)
root.iconphoto(True, PhotoImage(file=get_res_path('app_logo.png')))

song_info_label_frame = LabelFrame(root, text='Song Info' if not is_CN else '歌曲信息')
song_info_label_frame.pack(padx=10, pady=10, fill='both', expand=True)
song_title_text = StringVar(value='N/A')
song_artist_text = StringVar(value='')
title_label = Label(song_info_label_frame, textvariable=song_title_text)
title_label.pack(padx=10, pady=5)
artist_label = Label(song_info_label_frame, textvariable=song_artist_text)
artist_label.pack(padx=10, pady=5)

toggle_var = BooleanVar()
toggle_var.set(False)
toggle_button_text = StringVar(value='Enabled - Click to disable' if not is_CN else '已启用 - 点击以禁用')
toggle_var.trace_add('write', lambda *args: toggle_button_text.set(('Enabled - Click to disable' if not is_CN else '已启用 - 点击以禁用') if toggle_var.get() else ('Disabled - Click to enable' if not is_CN else '已禁用 - 点击以启用')))  # noqa
toggle_button = Button(root, textvariable=toggle_button_text, command=toggle, width=50)
toggle_button.pack(padx=10, pady=(10, 5))

startup_var = BooleanVar()
startup_var.set(os.path.isfile(startup_file_path))
startup_checkbox = Checkbutton(root, text='Start with Windows' if not is_CN else '开机自启', variable=startup_var, command=toggle_startup)
startup_checkbox.pack(padx=10, pady=5)

about_button = Button(root, text='About' if not is_CN else '关于', command=about, width=50)
about_button.pack(padx=10, pady=5)

github_button = Button(root, text='GitHub', command=lambda: webbrowser.open('https://github.com/aliencaocao/netease_cloudmusic_discord_rpc'), width=50)
github_button.pack(padx=10, pady=5)

minimize_button = Button(root, text='Minimize' if not is_CN else '最小化到托盘', command=hide_window, width=50)
minimize_button.pack(padx=10, pady=5)

quit_button = Button(root, text='Quit' if not is_CN else '退出', command=quit_app, width=50)
quit_button.pack(padx=10, pady=(5, 10))

root.protocol('WM_DELETE_WINDOW', hide_window)  # override close button to minimize to tray
root.after_idle(startup)
root.mainloop()
