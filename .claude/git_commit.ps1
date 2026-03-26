$ErrorActionPreference = "Continue"
$repo = "C:\Users\User\.claude\debug\.claude\worktrees\gracious-stonebraker"
$git = "C:\Program Files\Git\bin\git.exe"
Set-Location $repo

Write-Host "=== GIT STATUS ==="
& $git status

Write-Host "=== GIT ADD ==="
& $git add .

Write-Host "=== GIT COMMIT ==="
& $git commit -m "Add FOT PWA for Moskva 2 KK

FastAPI+SQLite backend, mobile PWA, PIN auth.
Railway deployment config.

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"

Write-Host "=== GIT PUSH ==="
& $git push origin claude/gracious-stonebraker

Write-Host "=== GIT LOG ==="
& $git log --oneline -3

Write-Host "=== Starting TCP listener ==="
$tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, 8000)
$tcp.Start()
Write-Host "Listening on port 8000"
while ($true) {
    $client = $tcp.AcceptTcpClient()
    $stream = $client.GetStream()
    $body = "Git commit done"
    $response = "HTTP/1.1 200 OK`r`nContent-Length: $($body.Length)`r`n`r`n$body"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($response)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Close()
    $client.Close()
}
