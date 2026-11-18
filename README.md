# 演示倒计时 2.0

适用于 Windows 的便携倒计时浮窗。支持 Microsoft PowerPoint 桌面版、WPS 演示的幻灯片放映，以及 WPS PDF 的全屏阅读。

## 使用方法

1. 在 Releases 下载 `presentation-countdown-2.0-windows.exe`，双击运行。
2. 右键点击浮窗，选择「设置」，输入倒计时秒数。
3. 在 PowerPoint 或 WPS 演示中开始放映。用 WPS 播放 PDF 时，按 F11 进入全屏。
4. 放映开始后浮窗自动倒计时，翻页时持续计时，归零后变红。退出放映后，浮窗显示最近保存的时长。

按住浮窗上的鼠标左键可拖动。工具会保存时长和位置，右键菜单可打开设置、查看关于或退出。

## 运行环境

Windows 10 或 11，使用系统的 .NET Framework 4.x。自动触发倒计时需要桌面版 PowerPoint 或 WPS Office。

## 从源码构建

在项目根目录用 Windows PowerShell 运行 `.\build.ps1`。脚本使用 Windows 的 .NET Framework C# 编译器，生成 `dist\presentation-countdown.exe`。

开发：曾峻峰（53106995@qq.com）
