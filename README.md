**USB GUID Retriever**  
*A simple PowerShell GUI tool to assist Windows users to retrieve USB drives' GUID info* 
![Screenshoot](https://raw.githubusercontent.com/yanzhou-ca/USBGUIDRetriever/refs/heads/main/Screenshot.png "USB GUID Retriever")
---

**Features**  
- 📌 Graphical interface for USB drive management  
- 🔍 Automatic detection of connected USB drives  
- 📋 Copy GUIDs to clipboard with one click  
- 🔄 Multi-drive support with dropdown selection  
- ⚡ Automatic loading of first detected drive's GUID  
- 📂 Activity logging to `ProgramData\USBGUIDRetriever\logs`  

---

**Requirements**  
- Windows 10/11 with PowerShell 5.1+  
- Administrator privileges  
- Execution Policy set to `RemoteSigned` or `Unrestricted`  

---

**Usage**  
1. Run as Administrator:  
   ```powershell
   .\USBGUIDRetriever.ps1
   ```
2. Insert USB drive(s)  
3. Interface components:  
   - **Dropdown**: Select between multiple USB drives  
   - **Get GUID**: Display selected drive's identifier  
   - **Copy**: Copy GUID to clipboard  
   - **Refresh**: Rescan for connected drives  

*Note: Drives connected before script launch will auto-load their GUID*

---

**Manual Verification**  
To confirm GUID values via PowerShell:  
```powershell
Get-CimInstance -ClassName Win32_Volume | 
    Where-Object { $_.DriveType -eq 2 } |
    Select-Object DriveLetter, DeviceID
```

---

**Logging**  
- Activity log: `%ProgramData%\USBGUIDManager\logs\activity.log`  
- System errors logged to Windows Event Viewer (Application log)

---

**License**  
MIT License - Free for personal/business use. No warranties provided.  

*Note: Requires .NET Framework 4.8 (included in Windows 10/11)*
