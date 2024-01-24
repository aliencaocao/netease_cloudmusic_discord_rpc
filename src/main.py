import ctypes
import gc
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
import pythoncom
import wmi
from PIL import Image
from pyMeow import close_process, get_module, get_process_name, open_process, pid_exists, pointer_chain, r_bytes, r_float64, r_uint
from pyncm import apis
from pypresence import DiscordNotFound, PipeClosed, Presence
from pystray import Icon as TrayIcon, Menu as TrayMenu, MenuItem as TrayItem
from win32api import GetFileVersionInfo, HIWORD, LOWORD

__version__ = '0.3.3'

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
    '2.10.5.3929': {'current': 0xA47548, 'song_array': 0xAF6FC8},
    '2.10.6.3993': {'current': 0xA65568, 'song_array': 0xB15654},
    '2.10.7.4239': {'current': 0xA66568, 'song_array': 0xB16974},
    '2.10.8.4337': {'current': 0xA74570, 'song_array': 0xB24F28},
    '2.10.10.4509': {'current': 0xA77580, 'song_array': 0xB282CC},
    '2.10.10.4689': {'current': 0xA79580, 'song_array': 0xB2AD10},
    '2.10.11.4930': {'current': 0xA7A580, 'song_array': 0xB2BCB0},
    '2.10.12.5241': {'current': 0xA7A580, 'song_array': 0xB2BCB0}, }
# '3.0.6.5811': {'current': 0x192B7F0, 'song_array': 0x0196DC38, 'song_array_offsets': [0x398, 0x0, 0x0, 0x8, 0x8, 0x50, 0xBA0]}, }  # TODO: song array offsets are different for every session, current and song_array stays same

interval = 1
windll = ctypes.windll.kernel32
is_CN = locale.windows_locale[windll.GetUserDefaultUILanguage()].startswith('zh_')

# regexes
re_song_id = re.compile(r'(\d+)')
logger.info(f"Netease Cloud Music Discord RPC v{__version__}\nRunning on Python {sys.version}\nSupporting NCM version: {', '.join(offsets.keys())}")


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
    try:
        presence.connect()
    except DiscordNotFound:
        connected_var.set(False)
        logger.warning('Discord not found.')
        messagebox.showerror('Discord not found', 'Could not detect a running Discord instance. Please make sure Discord is running and try again. Do not use BetterDiscord or other 3rd party clients.')
        return False
    except Exception as e:
        connected_var.set(False)
        logger.warning('Error while connecting to Discord:', e)
        return False
    else:
        connected_var.set(True)
        logger.info('Discord Connected.')
        return True


client_id = '1045242932128645180'
RPC = Presence(client_id)

first_run = True
pid = 0
version = ''
last_status = Status.changed
last_id = ''
last_float = 0.0
stop_variable = ThreadingEvent()

song_info_cache: Dict[str, SongInfo] = {}


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


def about():
    messagebox.showinfo('About', f"Netease Cloud Music Discord RPC v{__version__}\nPython {sys.version}\nSupporting NCM version: {', '.join(offsets.keys())}\nMaintainer: Billy Cao" if not is_CN else
    f"网易云音乐 Discord RPC v{__version__}\nPython版本 {sys.version}\n支持的网易云音乐版本: {', '.join(offsets.keys())}\n开发者: Billy Cao")


def quit_app(icon=None, item=None):
    stop_update()
    if icon: icon.stop()
    root.destroy()


def show_window(icon, item):
    icon.stop()
    root.after(0, root.deiconify())


def hide_window():
    global icon
    root.withdraw()
    if toggle_var.get():
        menu = TrayMenu(*[disable_item] + org_menu)
    else:
        menu = TrayMenu(*[enable_item] + org_menu)
    icon = TrayIcon("Netease Cloud Music Discord RPC", icon_image, "Netease Cloud Music Discord RPC", menu)
    icon.run()


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


def get_song_info(song_id: str) -> SongInfo:
    global song_info_cache
    if song_id not in song_info_cache:
        if not get_song_info_from_local(song_id):
            get_song_info_from_netease(song_id)
    return song_info_cache[song_id]


def find_process() -> Tuple[int, str]:
    pythoncom.CoInitialize()
    wmic = wmi.WMI()
    process_list = wmic.Win32_Process(name='cloudmusic.exe')
    process_list = [p for p in process_list if '--type=' not in p.CommandLine]
    if not process_list:
        return 0, ''
    if len(process_list) > 1:
        raise RuntimeError('Multiple candidate processes found!')
    process = process_list[0]

    ver_info = GetFileVersionInfo(process.ExecutablePath, '\\')
    ver = (f"{HIWORD(ver_info['FileVersionMS'])}.{LOWORD(ver_info['FileVersionMS'])}."
           f"{HIWORD(ver_info['FileVersionLS'])}.{LOWORD(ver_info['FileVersionLS'])}")

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
            stop_variable.set()
            raise UnsupportedVersionError(f"This version is not supported yet: {version}.\nSupported version: {', '.join(offsets.keys())}" if not is_CN else f"目前不支持此网易云音乐版本: {version}。\n支持的版本: {', '.join(offsets.keys())}")

        if first_run:
            logger.info(f'Found process: {pid}')
            first_run = False

        process = open_process(pid)
        module_base = get_module(process, 'cloudmusic.dll')['base']

        current_float = r_float64(process, module_base + offsets[version]['current'])
        current_pystr = sec_to_str(current_float)
        if version.startswith('2.'):
            songid_array = r_uint(process, module_base + offsets[version]['song_array'])
            song_id = (r_bytes(process, songid_array, 0x14).decode('utf-16').split('_')[0])  # Song ID can be shorter than 10 digits.
        elif version.startswith('3.'):
            songid_array = pointer_chain(process, module_base + offsets[version]['song_array'], offsets[version]['song_array_offsets'])
            song_id = r_bytes(process, songid_array, 0x14)
            song_id = song_id.decode('utf-16').replace('\x00', '').split('_')[0]
        else:
            raise RuntimeError(f'Unknown version: {version}')

        close_process(process)

        if not re_song_id.match(song_id):
            # Song ID is not ready yet.
            return

        # Interval should fall in (interval - 0.2, interval + 0.2)
        status = (Status.playing if song_id == last_id and abs(current_float - last_float - interval) < 0.1
                  else Status.paused if song_id == last_id and current_float == last_float
        else Status.changed)

        if status == Status.playing:
            if last_status != Status.paused:  # Nothing changed
                last_float = current_float
                last_status = Status.playing
                return
        elif status == Status.paused:
            if last_status == Status.paused:  # Nothing changed
                return
            else:
                logger.debug('Paused')

        song_info = get_song_info(song_id)

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
                       buttons=[{'label': 'Listen on Netease',
                                 'url': f'https://music.163.com/#/song?id={song_id}'}]
                       )
        except PipeClosed:
            logger.info('Reconnecting to Discord...')
            if connect_discord(RPC):
                connected_var.set(True)
            else:
                connected_var.set(False)
        except Exception as e:
            logger.error('Error while updating to Discord:')
            logger.exception(e)
            pass

        last_id = song_id
        last_float = current_float
        last_status = status

        if status != Status.paused:
            logger.debug(f"{song_info['title']} - {song_info['artist']}, {current_pystr}")
    except UnsupportedVersionError as e:
        messagebox.showerror('不支持的网易云音乐版本', str(e))
        toggle_var.set(False)
        stop_variable.set()
        raise e
    except Exception as e:
        logger.error('Error while updating song info:')
        logger.exception(e)
    gc.collect()


def startup():
    global timer
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
    stop_variable.set()
    if 'timer' in globals():
        timer.stop()
    RPC.close()


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

toggle_var = BooleanVar()
toggle_var.set(False)
toggle_button_text = StringVar(value='Enabled - Click to disable' if not is_CN else '已启用 - 点击以禁用')
toggle_var.trace_add('write', lambda *args: toggle_button_text.set(('Enabled - Click to disable' if not is_CN else '已启用 - 点击以禁用') if toggle_var.get() else ('Disabled - Click to enable' if not is_CN else '已禁用 - 点击以启用')))  # noqa
connected_var = BooleanVar()
connected_var.set(False)

toggle_button = Button(root, textvariable=toggle_button_text, command=toggle, width=80)
toggle_button.pack(padx=10, pady=(10, 5))

about_button = Button(root, text='About' if not is_CN else '关于', command=about, width=80)
about_button.pack(padx=10, pady=5)

github_button = Button(root, text='GitHub', command=lambda: webbrowser.open('https://github.com/aliencaocao/netease_cloudmusic_discord_rpc'), width=80)
github_button.pack(padx=10, pady=5)

minimize_button = Button(root, text='Minimize' if not is_CN else '最小化到托盘', command=hide_window, width=80)
minimize_button.pack(padx=10, pady=5)

quit_button = Button(root, text='Quit' if not is_CN else '退出', command=quit_app, width=80)
quit_button.pack(padx=10, pady=(5, 10))

root.protocol('WM_DELETE_WINDOW', hide_window)
root.after_idle(startup)
root.mainloop()
