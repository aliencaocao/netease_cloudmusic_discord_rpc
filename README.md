# 网易云音乐 Discord Rich Presence (RPC)

## 介绍 About
支持同步歌曲，歌手和目前歌曲的播放时长。歌曲的总时长暂不支持显示。

Supports synchronizing song, artist and current song's playing time. Total duration of the song is not supported yet.

纯Python写成，支持最新版网易云音乐，目前只支持Windows客户端。

Written in pure Python, supports latest version of NetEase Cloud Music. Currently only supports the Windows client.

目前只支持网易云音乐2.10.6 build 200601，还会继续支持未来的新版本。

Currently ONLY supports NetEase Cloudmusic Windows 2.10.6 build 200601. Support for future versions will be added once they are released.

旧版本(2.10.2及以下)可以使用这个项目：https://github.com/Kxnrl/NetEase-Cloud-Music-DiscordRPC

For older versions (2.10.2 and below), you can checkout this project: https://github.com/Kxnrl/NetEase-Cloud-Music-DiscordRPC

效果图 / Demo

![demo](demo.png)


## 使用方法 Usage
从Release页下载可执行文件(exe)，然后运行exe并保持运行。

程序不会显示任何窗口，而是在后台保持运行。

程序必须以UAC模式运行（管理员权限）才能正常工作。

Download the executable binary (exe) from the Release page, Keep the exe running.

The program will not show any window and will run in the background.

The program must be opened in UAC mode (Administrator permission) to work properly.

# 构建 Building
你需要 / You need:
- Python 3.6+
- `pip install -r requirements.txt`

运行build.txt里面的命令即可。

Run the commands in build.txt.

Inspired by https://github.com/Kxnrl/NetEase-Cloud-Music-DiscordRPC