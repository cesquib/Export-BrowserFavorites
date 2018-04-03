Try {
    "This is a test" | Out-File -FilePath 'C:\windows\system32\test.html' -Force -Encoding utf8
} Catch {
    Write-Error -Message "Failed to write to 'C:\windows\system32\test.html'. Please check permissions."
    Exit
}


