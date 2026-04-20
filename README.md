# Exchange Queue Manager PRO

A high-performance, GUI-based management tool for Microsoft Exchange Server (2013, 2016, 2019). Designed for systems administrators who need a fast, automated way to monitor and clear mail queues without using the command line.

## Key Features
- **Automated Loading**: Fetches all queues immediately upon startup.
- **Smart Selection**: Automatically loads messages for a queue when you select it in the list.
- **Live Filtering**: Filter queues by name or status (Active, Retry, Suspended) in real-time.
- **Bulk Actions**: Clear specific messages, entire selected queues, or all filtered queues at once.
- **Multi-Threaded**: Background engine keeps the UI responsive even during large data fetches.
- **Cross-Version**: Native support for PowerShell 5.1 and 7.

## Installation
1. Download the files `Queues.ps1` and `Launch.vbs` to a folder on your Exchange server.
2. Ensure you have the **Exchange Management Tools** installed locally.

## Running the Tool

### The "Invisible" Shortcut (Recommended)
To run the tool without a PowerShell console window appearing:
1. Right-click on **Launch.vbs** and select **Create Shortcut**.
2. Move this shortcut to your Desktop.
3. (Optional) Right-click the shortcut > Properties > **Change Icon**. 
   - Browse to `C:\Windows\System32\shell32.dll` and pick a network/server icon.

### Manual Run
Alternatively, right-click `Queues.ps1` and select **Run with PowerShell**.

## Requirements
- Windows Server with Microsoft Exchange installed.
- PowerShell Execution Policy set to `Bypass` or `RemoteSigned`.

## License
This project is licensed under the MIT License - feel free to use and modify for your organization.
<img width="1179" height="708" alt="image" src="https://github.com/user-attachments/assets/ceb40998-a7b9-4ab1-925e-8d2d22fb729e" />
