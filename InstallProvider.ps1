$p = 'C:\Program Files\WindowsPowerShell\Modules\WUAProvider'
try {
    if (-not(Test-Path -Path $p -PathType Container) ) {
        $null = mkdir -Path $p -EA Stop
    }
    'd1','m1' | ForEach-Object {
        Copy-Item -Path "WUAProvider.ps$($_)" -Destination $p -EA Stop
    }
} catch {
    throw $_
}