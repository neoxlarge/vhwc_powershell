


function test-url ($url) {

    $response = Invoke-WebRequest -uri $url

    if ($response.StatusCode -eq 200) {
        Write-Host (get-date)
        Write-Host "$url �������T���J�I"
    } else {
        Write-warning (get-date)
        Write-Warning "$url �������J���ѡC���A�X: $($response.StatusCode)"

        [console]::beep(1000, 3000)

        throw "$url �������J���ѡC���A�X: $($response.StatusCode)"

    }

}
$u =+ 1
do {

    $urls = "http://eip.vghtc.gov.tw", "https://mail.vghtc.gov.tw", "http://edsap.edoc.vghtc.gov.tw","http://edsap1.edoc.vghtc.gov.tw"

    foreach ($u in $urls) {

        test-url -url $u

        Start-Sleep -Seconds 60

    }
    
    Start-Sleep -Seconds 300

} while (
    $true
)

