[简体中文](README_zh.md) | English
# 🚀 Auto Check-in System (某安 Auto Attendance)

A Windows desktop automation tool built with Python + Tkinter + UI Automation. It is designed to handle repetitive daily mini-program check-in tasks, supporting scheduled execution, WeChat notifications, and startup with Windows.

## ✨ Core Features

- 🤖 **Fully Automated Check-in**: Automatically wakes the screen, unlocks the system, launches WeChat mini-programs, and completes the check-in process.
- ⏰ **Smart Scheduled Tasks**: Supports custom check-in times with a built-in anti-sleep mechanism (allows screen off but prevents system hibernation).
- 📱 **WeChat Status Notification**: Automatically sends the check-in result to "File Transfer Helper" upon completion, with an option to close WeChat afterward.
- 💻 **Startup Support**: One-click configuration of Windows Task Scheduler to run automatically after user login.
- 🛡️ **Robustness Optimizations**:
    - Built-in DPI scaling support for high-resolution screens.
    - Retry mechanisms for network disconnections or lag, and automatic handling of conflicting pop-ups.
    - Support for a "Missed Check-in Grace Period" recovery function.

## ⚠️ Pre-requisites & Important Notes

- **WeChat Environment**: Ensure the PC version of WeChat is logged in and running in the background.
- **Shortcut**: A shortcut to the target mini-program must exist on the Desktop (Default keyword: "中南林业科技大学学生工作部". This can be modified in the code or configuration).
- **Lock Screen Settings**: It is recommended to disable the Windows login password or enable Auto-login. Otherwise, the system cannot automatically enter the Desktop when locked.
- **Administrator Privileges**: When using "Startup with Windows" or "Anti-Sleep" features for the first time, please **Run as Administrator**.

## 📥 Installation & Execution

1. Download the latest `.exe` installer package.
2. Extract the files, right-click the main program, and select **"Run as Administrator"**.
3. Set the desired check-in time in the interface and click **"Start Scheduled Check-in"**.
