![invoke-wordthief-logo](https://user-images.githubusercontent.com/53918129/88349182-ff761900-cd57-11ea-8c44-816844ed76d8.png)
This script runs multithreading module that connects to a remote TCP server,
monitors active (opened) Microsoft Word documents (.doc,.docx,etc') and extracting
their text using Word application's COM Object.
The script adds HKCU registry (no admin needed) Run key, so this script runs persistently
## Special Thanks  
*  [.NET Core documentation](https://docs.microsoft.com/en-us/dotnet/api/?view=netcore-3.1) - helped me to figure out all about COM Objects and how to play with them
* Multi-threaded TCP server implementation in Python - https://www.geeksforgeeks.org/socket-programming-multi-threading-python/
* Persistence: Registry Run Keys - https://pentestlab.blog/2019/10/01/persistence-registry-run-keys/
* Every good guy that posts an informative answer to stupid questions in Stack Overflow
## How To Use
```
# Run attacker's log server with Python 3
python ./logger.py -h

# Run Powershell in Windows machine, show module help info and examples
PS> Import-Module .\Invoke-WordThief.ps1
PS> help Invoke-WordThief
PS> help Invoke-WordThief -Examples
```

## Research Overview
At the beginning, I wasn't sure how I should extract text from active Word processes. I searched a bit online but I figured out quickly it isn't much of a legit action to do, so I didn't find much.<br/>I started by analysing WINWORD.EXE processes with [SysInternals](https://docs.microsoft.com/en-us/sysinternals/) tools like ProcMon and ProcessExplorer, but those didn't fit in this specific task.<br/>
I kept digging the internet until I encountered [COM Objects](https://docs.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal).<br/>With some reading, I figured out some core methods I can use to get an handle of active documents in Microsoft Office, for example: [GetActiveObject()](https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.getactiveobject), [Document Interface in Office API](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.tools.word.document).<br/>
From there, the development went straight forward. I decided to get along with victim's environment (living-of-the-land style), so I built a Powershell [multi job](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/start-job) script, means a background job for each opened document. I chose Powershell because every Windows 10 has the engine built-in, and accessing .NET API is pretty comfortable.<br/>Alongside the Powershell tool, there is a Python listener that should receive the text and piping it to local files in the attacker's machine.<br/>This tool was fun to build and I hope I'll find the time to create more tools like it because I can find the use in many situations in Red Team operations.
