function base64-decode {
    param (
        [String]$B64
    )

    return [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($B64))
}