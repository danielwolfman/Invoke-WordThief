# Word-Monitor

This script runs a key logger that records keypresses while user has opened Microsoft Word. After the key press has been captured, its being sent to attacker server through TCP session. This script has also persistency function, spawning a batch file in user's startup folder.

## Why PowerShell?<img src="https://1.bp.blogspot.com/-trcervxbi1c/Wqvxd3tIdAI/AAAAAAAAK3s/x4CX7QkRAbY1OdnFediDFeR7eG9M_R-iwCLcBGAs/s1600/PowerShell_5.0_icon.png" width="30" height="30" />

PowerShell framwork is in every Windows operation system since 2006.
PowerShell is usually a task automation and configuration management framework from [Microsoft](https://en.wikipedia.org/wiki/Microsoft "Microsoft"), but attackers can use it to execute code conveniently on any Windows machine.
I also used various of [.NET Core](https://docs.microsoft.com/en-us/dotnet/api/?view=netcore-3.1) classes so I can use their useful functions that helps me build TCP socket, key press capturing, etc'.

## Microsoft Word

The goal was to extract text from opened or newly created Word documents.
After researching about WINWORD.EXE process, I figured out 2 things:
 - If you double click a .doc/.docx file, WINWORD.exe process is running with an <ins>argument of the full path of the file</ins>. So what I could do is just monitoring new processes and catching the argument list of new WINWORD.exe processes and binary reading the file and sending in through my TCP socket. (pretty simple, isn't it?)
 - If you open a file after you started Word, that's where it's getting tricky.

