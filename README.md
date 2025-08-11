# powershell-archive-object-extractor
A PowerShell script and Windows executable providing an alternative method to extract `.zip` and `.tar` files on Windows. It improves on the built-in File Explorer extractor tool—which was malfunctioning and unusable for sporadic corporate users—and offers a reliable alternative to limited-functionality third-party tools like WinZip.

---

## Background and Inspiration

During my time working helpdesk support, sporadic corporate users experienced constant failures with the native Windows File Explorer extractor tool. The tool was malfunctioning and often unusable, causing frustrating delays in workflows that depended on quick archive extraction.

While exploring solutions, I found that third-party tools such as WinZip offered limited features without paid licenses after their trial periods, making them impractical for long-term use.

To solve this, I created this PowerShell-based archive extractor as a free, dependable, and user-friendly alternative that reliably handles multiple archive formats without the drawbacks of both the malfunctioning native tool and restrictive third-party software.

---

## Features

- Extract `.zip`, `.tar`, `.tar.gz`, `.tgz`, and `.tar.bz2` archives  
- Supports both GUI and command-line modes  
- Displays completion messages with extraction paths  
- Logs errors to a local log file for troubleshooting  
- Portable executable available — no PowerShell setup required  

---

## Requirements

- Windows 10 or later  
- PowerShell 5.1 or higher (usually pre-installed on Windows 10+)  
- `tar.exe` included in Windows 10+ for `.tar` and `.tar.gz` extraction  
- .NET Framework (built into Windows)  

---

## Installation

Before starting, please ensure:  

- Your Windows system is fully updated with the latest patches and updates.  
- PowerShell version 5.1 or above is installed. To check your PowerShell version, open PowerShell and run:

  ```powershell
  $PSVersionTable.PSVersion
