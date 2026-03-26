Write-Host "Starting test server..."
$tcp = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Any, 8000)
$tcp.Start()
Write-Host "Listening on 8000"
while ($true) {
    $client = $tcp.AcceptTcpClient()
    $stream = $client.GetStream()
    $body = "OK"
    $resp = "HTTP/1.1 200 OK`r`nContent-Length: 2`r`n`r`nOK"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($resp)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Close()
    $client.Close()
}
