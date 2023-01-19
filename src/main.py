import gc
import time
from threading import Event, Thread
import ctypes
from ctypes import wintypes
import wmi
import pythoncom
from win32com.client import Dispatch
from pypresence import Presence
from pyMeow import open_process, get_module, r_float64, close_process, get_module
from typing import Tuple

__version__ = '0.1.2'
supported_cloudmusic_version = '2.10.6.3993'
current_offset = 0xA65568
maxlen_offset = 0xB16438  # TODO: does not work
print(f'网易云音乐Discord RPC v{__version__}，支持网易云音乐版本：{supported_cloudmusic_version}')

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


class RepeatedTimer:
    def __init__(self, interval, function):
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


def get_title(pid) -> Tuple[str, str]:
    title = ''
    artist = ''

    @WNDENUMPROC
    def enum_proc(hWnd, lParam):
        nonlocal title
        nonlocal artist

        if user32.IsWindowVisible(hWnd):
            _pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hWnd, ctypes.byref(_pid))
            length = user32.GetWindowTextLengthW(hWnd) + 1
            buf = ctypes.create_unicode_buffer(length)
            user32.GetWindowTextW(hWnd, buf, length)

            if _pid.value == pid: 
                title = buf.value
                if ' - ' in title:
                    title, artist = title.split(' - ')
        return True

    user32.EnumWindows(enum_proc, 0)
    return (title, artist)


def sec_to_str(sec) -> str:
    m, s = divmod(sec, 60)
    return f'{int(m):02d}:{int(s):02d}'


client_id = '1065646978672902144'
RPC = Presence(client_id)
RPC.connect()
print('RPC Launched\nThe following info will only be printed once for confirmation. They will continue to be updated to Discord.')
start_time = time.time()
first_run = True


def update():
    global first_run
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
        info = ver_parser.GetFileVersion(process.ExecutablePath)
        if info != supported_cloudmusic_version: raise RuntimeError(f'This version is not supported yet: {info}. Supported version: {supported_cloudmusic_version}')
        pid = process.ole_object.ProcessId
        if first_run: print(f'Found process: {pid}')
    (title, artist) = get_title(pid)
    if not title: title = 'Unknown'
    process = open_process(pid)
    base_address = get_module(process, 'cloudmusic.dll')['base']
    current = r_float64(process, base_address + current_offset)
    current_s = sec_to_str(current)
    # maxlen = r_float64(process, base_address + maxlen_offset)  # not working now
    close_process(process)
    RPC.update(pid=pid, details=f'{title}', state=f'{artist}', large_image='logo', large_text='Netease Cloud Music', start=int(time.time() - current))
    if first_run: print(f'{title} - {artist}, current: {current_s}')
    first_run = False
    gc.collect()
    pythoncom.CoUninitialize()


RepeatedTimer(1, update)  # calls the update function every second, ignore how long the actual update takes
