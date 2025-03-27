# Toggle Taskbar Auto-hide

![ToggleTaskbarAutohide Logo](demo.gif)

**A convenient shortcut to directly toggle the "Automatically hide the taskbar in desktop mode" option, with optional system tray shortcut.*

# Implementation Details
I originally implemented this as a .bat script like so
``@echo off
powershell -command "&{$p='HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3';$v=(Get-ItemProperty -Path $p).Settings;if($v[8] -eq 2){$v[8]=3;}else{$v[8]=2;}&Set-ItemProperty -Path $p -Name Settings -Value $v;&Stop-Process -f -ProcessName explorer}"``
But the problem with this was that a command prompt window would pop up for a split second, which is visually annoying. so it was required to rewrite this as an application that used system calls directly to do it "silently". 


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
