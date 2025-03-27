# Toggle Taskbar Auto-hide

![ToggleTaskbarAutohide Logo](demo.gif)

**A convenient shortcut to directly toggle the "Automatically hide the taskbar in desktop mode" option, with optional system tray shortcut.*

# Implementation Details
I originally implemented this as a .bat script like so
``@echo off
powershell -command "&{$p='HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3';$v=(Get-ItemProperty -Path $p).Settings;if($v[8] -eq 2){$v[8]=3;}else{$v[8]=2;}&Set-ItemProperty -Path $p -Name Settings -Value $v;&Stop-Process -f -ProcessName explorer}"``
But the problem with this was that a command prompt window would pop up for a split second, which is visually annoying. so it was required to rewrite this as an invisible application with direct system calls. The StuckRects3 can be documented as follows:
```
    ╔═══════════════════════════════════════════════════════════════════════════════╗
    ║ StuckRects3 "Settings" Binary Structure Map (64-byte serialized configuration)║
    ╠═══════════════════════════════════════════════════════════════════════════════╣
    ║ Offset │ Size │ Purpose                                                       ║
    ╠════════╪══════╪═══════════════════════════════════════════════════════════════╣
    ║ 0x00   │ 4    │ Structure version identifier (typically 0x30,0x00,0x00,0x00)  ║
    ║ 0x04   │ 4    │ Configuration bitflags:                                       ║
    ║        │      │ • Bit 0: Taskbar position (0=bottom, 1=top, 2=left, 3=right)  ║
    ║        │      │ • Bits 1-31: Reserved for internal Windows use                ║
    ║ 0x08   │ 1    │ ► Visibility control flag:                                    ║
    ║        │      │ • 0x02 = Always visible (standard configuration)              ║
    ║        │      │ • 0x03 = Auto-hide enabled                                    ║
    ║ 0x09   │ 3    │ Reserved for future use                                       ║
    ║ 0x0C   │ 16   │ Taskbar position/dimension information                        ║
    ║ 0x1C   │ 36   │ Additional configuration data                                 ║
    ╚════════╧══════╧═══════════════════════════════════════════════════════════════╝
```
The rest of the code deals with the tray icon and reopening explorer windows state since they have to be closed and reopened for the setting change to take effect. 


[Download Latest Release](https://github.com/FreelanceProgrammingServices/ToggleTaskbarAutohide/releases/latest)

</div>

## Installation

1. Download the latest `ToggleTaskbarAutohide.exe` from the [releases page](https://github.com/yourusername/ToggleTaskbarAutohide/releases/latest) 
  There are 2 versions. 1 version directly toggles the shortcut, and the 2nd provides a persistent system tray shortcut.
2. Place the executable in any location you prefer.
3. Run the application directly - no installation needed.

## Support
Tested on:
  Windows 8, 10, and 11.


## License

This project is released under Public Domain.

---

</div>
